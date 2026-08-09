#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Aggregazione indici clinici da piu articoli + motore di valutazione della pericolosita.
Articolo 1: indici cardiologici.  Articolo 2: indici antropometrici."""

import os

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.comments import Comment
from openpyxl.formatting.rule import FormulaRule

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                   "Indici_Clinici_Analisi_Rischio.xlsx")
FONT = "Arial"
BLUE, BLACK, GREEN, YEL = "0000FF", "000000", "008000", "FFFF00"
HDR_FILL = PatternFill("solid", fgColor="1F3864")
SUB_FILL = PatternFill("solid", fgColor="D9E2F3")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
BAD_FILL = PatternFill("solid", fgColor="F8CBAD")
ERR_FILL = PatternFill("solid", fgColor="FCE4D6")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)

A1 = "Art.1 - Indici cardiologici"
A2 = "Art.2 - Indici antropometrici"
A3 = "Art.3 - Profilo marziale"
A4 = "Art.4 - Indici infiammatori"

# Pesi molecolari usati per convertire il pannello marziale dalle unita SI
# a quelle convenzionali. Documentati perche il rapporto transferrina/log(ferritina)
# ne dipende: vedi foglio 08.
MW_TRF = 79570      # transferrina, g/mol
MW_FER = 450000     # ferritina, g/mol
K_TRF_GL = MW_TRF / 1e6        # uM -> g/L
K_FER_UGL = MW_FER / 1e6 / 1e3  # pM -> ug/L
K_FE_UGDL = 5.585              # uM -> ug/dL (ferro)


def hdr(ws, row, labels, start=1, h=32):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=lab)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = h


def title(ws, text, sub=None):
    c = ws.cell(row=1, column=1, value=text)
    c.font = Font(name=FONT, size=14, bold=True, color="1F3864")
    if sub:
        c2 = ws.cell(row=2, column=1, value=sub)
        c2.font = Font(name=FONT, size=9, italic=True, color="595959")


def setw(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w


def rowh(nota, width=78, lo=30, hi=80):
    return max(lo, min(hi, 13 * (len(nota) // width + 1)))


wb = Workbook()

# =====================================================================
# 00 LEGENDA
# =====================================================================
ws = wb.active
ws.title = "00 Legenda"
setw(ws, {"A": 32, "B": 100})
title(ws, "Indici clinici - aggregazione multi-articolo, calcolo e valutazione della pericolosita",
      "Fonti integrate: Art.1 cardiologici, Art.2 antropometrici, Art.3 profilo marziale, Art.4 indici infiammatori. Stesso soggetto dimostrativo.")

LEG = [
    ("", ""),
    ("SCOPO", "Raccogliere in un unico modello formule e indici estratti dagli articoli forniti, ricalcolarli in "
              "modo trasparente dai dati grezzi del soggetto, verificare la riproducibilita dei valori pubblicati "
              "e produrre una valutazione strutturata del livello di pericolosita."),
    ("", ""),
    ("AVVERTENZA", "Strumento didattico e di sintesi documentale. NON e un dispositivo medico e non produce diagnosi. "
                   "'Gennarino Parsifallo' e dichiarato dalla fonte stessa come caso dimostrativo, non un paziente "
                   "reale. Ogni decisione clinica compete al medico."),
    ("", ""),
    ("COME SI USA", "1) Apri '01 Input' e modifica SOLO le celle blu su sfondo giallo.\n"
                    "2) Tutti gli altri fogli si ricalcolano da soli.\n"
                    "3) '02 Indici' contiene i 59 indici delle quattro fonti, ricalcolati e confrontati con il referto.\n"
                    "4) '06 Rischio' produce il punteggio complessivo di pericolosita.\n"
                    "5) '08 Controlli' elenca cosa e riproducibile e cosa no: leggilo prima di fidarti dei numeri."),
    ("", ""),
    ("CONVENZIONI COLORE", ""),
    ("  Testo blu su giallo", "Dato inserito a mano. E l'unica cosa da modificare."),
    ("  Testo nero", "Risultato calcolato nel foglio stesso."),
    ("  Testo verde", "Valore richiamato da un altro foglio."),
    ("  Riga arancione", "Indice fuori intervallo nella direzione critica."),
    ("  Riga gialla", "Indice dentro l'intervallo ma in zona di guardia."),
    ("", ""),
    ("FOGLI", ""),
    ("  01 Input", "Dati grezzi del soggetto, laboratorio e fattori anamnestici. Unico foglio da compilare."),
    ("  02 Indici", "I 59 indici delle quattro fonti: formula, valore ricalcolato, scostamento dal referto, intervallo "
                    "di riferimento, stato, peso clinico e punteggio."),
    ("  03 Formule", "Dizionario delle formule con derivazione e verifica numerica."),
    ("  04 Zone FC", "Zone di frequenza cardiaca ricavate dalla frequenza massima."),
    ("  05 Evidenze", "Letteratura citata dalle due fonti con PMID e applicabilita al caso."),
    ("  06 Rischio", "Motore di calcolo: punteggio strumentale per fonte, fattori anamnestici, rischio relativo "
                     "epidemiologico e indice composito di pericolosita."),
    ("  07 Scenari", "Simulazioni what-if sugli interventi discussi dalle fonti."),
    ("  08 Controlli", "Verifica di riproducibilita: quali valori pubblicati si riottengono dai dati grezzi e quali no."),
    ("  09 Conversioni", "Il profilo marziale tradotto dalle unita SI a quelle di un referto di laboratorio ordinario."),
    ("", ""),
    ("NOTA IMPORTANTE", "Il foglio '08 Controlli' registra gli scostamenti fra i valori pubblicati e il ricalcolo. "
                        "I due piu rilevanti: un BMI che non discende dal peso e dall'altezza dichiarati e che si "
                        "propaga su tre indici derivati, e una diagnosi di sarcopenia che poggia su una correzione "
                        "di cinque millimetri la cui formula non e pubblicata. Il modello mantiene sia il valore "
                        "pubblicato sia quello ricalcolato, senza sceglierne uno."),
    ("", ""),
    ("AGGIUNGERE UN ARTICOLO", "Aggiungi le righe in fondo alle tabelle di '02 Indici', '03 Formule' e '05 Evidenze' "
                               "compilando la colonna 'Fonte' con l'etichetta del nuovo articolo. Se introduce dati "
                               "grezzi nuovi, inseriscili prima in '01 Input'. I conteggi di '06 Rischio' usano "
                               "COUNTIFS sull'etichetta della fonte, quindi si aggiornano da soli."),
]
r = 4
for a, b in LEG:
    ca, cb = ws.cell(row=r, column=1, value=a), ws.cell(row=r, column=2, value=b)
    ca.font = Font(name=FONT, size=10, bold=bool(a and not a.startswith("  ")),
                   color="C00000" if a in ("AVVERTENZA", "NOTA IMPORTANTE") else "000000")
    cb.font = Font(name=FONT, size=10)
    cb.alignment = Alignment(wrap_text=True, vertical="top")
    ca.alignment = Alignment(vertical="top")
    if b:
        ws.row_dimensions[r].height = rowh(b, 100, 15, 62)
    r += 1

# =====================================================================
# 01 INPUT
# =====================================================================
ws = wb.create_sheet("01 Input")
setw(ws, {"A": 46, "B": 16, "C": 14, "D": 74})
title(ws, "01 - Dati grezzi del soggetto",
      "Modifica solo le celle blu su sfondo giallo. Tutto il resto del workbook si ricalcola.")
hdr(ws, 4, ["Parametro", "Valore", "Unita", "Fonte / Nota"])

INPUTS = [
    ("ANAGRAFICA", None, None, None),
    ("Identificativo soggetto", "Gennarino Parsifallo", "", "Caso dimostrativo dichiarato dalle fonti"),
    ("Eta", 50, "anni", A1 + " / " + A2),
    ("Sesso", "M", "M/F", "Determina il modello di FC massima, la stima del grasso corporeo e la formula del VAI"),
    ("", None, None, None),
    ("MISURE ANTROPOMETRICHE", None, None, None),
    ("Peso", 74.0, "kg", A2),
    ("Altezza", 174, "cm", A2),
    ("Circonferenza vita", 97, "cm", A2 + " - a fine espirazione, a meta fra ultima costa e cresta iliaca. Alimenta nove dei dodici indici antropometrici"),
    ("Circonferenza fianchi", 100, "cm", A2),
    ("Circonferenza polpaccio", 34, "cm", A2 + " - surrogato validato della massa muscolare appendicolare (PMID 30312372)"),
    ("Circonferenza polso", 18, "cm", A2 + " - misura stabile per tutta la vita adulta: e il riferimento fisso del morfotipo"),
    ("Peso all'eta di 30 anni", 68.0, "kg", A2 + " - serve a calcolare l'incremento ponderale ventennale"),
    ("Superficie corporea (BSA)", 1.9, "m2", A1 + " - riportata come 'circa 1,9 m2'; denominatore degli indici cardiologici indicizzati"),
    ("", None, None, None),
    ("MISURE CARDIOVASCOLARI", None, None, None),
    ("Pressione sistolica (SBP)", 128, "mmHg", A1),
    ("Pressione diastolica (DBP)", 84, "mmHg", A1),
    ("Frequenza cardiaca a riposo (FC)", 82, "bpm", A1 + " - unico indice cardiologico fuori range del referto"),
    ("Gittata sistolica (SV)", 66, "mL", A1),
    ("Frequenza cardiaca massima da referto", 175, "bpm", A1 + " - modello logistico sesso-specifico (PMID 12752560 / 11722475); i coefficienti non sono pubblicati, quindi il valore si inserisce a mano"),
    ("Pressione venosa centrale assunta (CVP)", 5, "mmHg", "ASSUNZIONE: non riportata dalle fonti. Serve solo alle resistenze vascolari. Con CVP=5 il calcolo restituisce i 2,7 kU*m2 del referto"),
    ("", None, None, None),
    ("LABORATORIO", None, None, None),
    ("Colesterolo HDL", 41, "mg/dL", A2 + " - preso da solo nessun laboratorio lo segnalerebbe. Alimenta CMI e VAI"),
    ("Trigliceridi", 148, "mg/dL", A2 + " - idem. Alimenta LAP, CMI e VAI"),
    ("", None, None, None),
    ("PROFILO MARZIALE (valori misurati)", None, None, None),
    ("Sideremia", 13.6, "uM", A3 + " - ritmo circadiano marcato: il prelievo va fatto al mattino a digiuno, altrimenti la saturazione risulta abbassata di diversi punti senza che le riserve siano cambiate"),
    ("Transferrina", 39.2, "uM", A3 + " - il fegato ne produce di piu proprio quando il ferro scarseggia. Da qui derivano TIBC e UIBC"),
    ("Ferritina", 160.9, "pM", A3 + " - proteina di fase acuta positiva: sale con l'infiammazione a prescindere da quanto ferro contenga"),
    ("Fattore di aggiustamento BRINDA - sideremia", 1.2574, "-", "DEDOTTO: rapporto fra il valore aggiustato e quello grezzo pubblicati (17,1/13,6). I marcatori di infiammazione su cui si basa la correzione non sono riportati dall'articolo"),
    ("Fattore di aggiustamento BRINDA - ferritina", 0.7502, "-", "DEDOTTO: rapporto fra il valore aggiustato e quello grezzo pubblicati (120,7/160,9). Corrisponde a una riduzione del 25% esatto"),
    ("", None, None, None),
    ("INFIAMMAZIONE - misure dirette", None, None, None),
    ("Proteina C reattiva (PCR)", 0.3, "mg/dL", A4 + " - proteina di fase acuta: sale molto e in fretta nell'infiammazione acuta, ed e per questo un cattivo strumento per quella cronica di basso grado"),
    ("VES a 1 ora", 14, "mm", A4 + " - misura una proprieta reologica del sangue che cambia nell'arco di giorni"),
    ("Omocisteina", 14.6, "uM", A4),
    ("Uricemia", 6.7, "mg/dL", A4),
    ("", None, None, None),
    ("EMOCROMO - valori ricostruiti per inversione degli indici", None, None, None),
    ("Piastrine", 272.0, "10^9/L", "DEDOTTO da AISI/SIRI. L'articolo non pubblica l'emocromo assoluto: i valori di questa sezione sono stati ricavati invertendo gli indici derivati"),
    ("Linfociti", 1.7407, "10^9/L", "DEDOTTO da (AISI/SII)/MLR"),
    ("Neutrofili", 5.8684, "10^9/L", "DEDOTTO da SIRI x linfociti / monociti. In percentuale l'articolo cita 69%, con limite a 70"),
    ("Monociti", 0.6615, "10^9/L", "DEDOTTO da AISI/SII. Corrisponde a circa il 7,8% della formula leucocitaria"),
    ("Volume piastrinico medio (MPV)", 11.906, "fL", "DEDOTTO da (MPV/linfociti) x linfociti"),
    ("Ampiezza di distribuzione eritrocitaria (RDW)", 14.5, "%", "ASSUNTO: il sistema e sottodeterminato. RDW/Piastrine e pubblicato con due sole cifre decimali (0,05), compatibile con un RDW fra 12,2 e 15,0. Scelto 14,5 perche coerente con l'anemia lieve descritta dalla serie"),
    ("Emoglobina", 13.9, "g/dL", "ASSUNTO: discende dall'RDW scelto tramite il rapporto Emoglobina/RDW pubblicato (0,96)"),
    ("", None, None, None),
    ("VALORI NON RICALCOLABILI (inseriti dal referto)", None, None, None),
    ("ABSI z-score da referto", 1.52, "z", A2 + " - il valore grezzo di ABSI e ricalcolabile, ma lo z-score richiede le tabelle NHANES per eta e sesso, che non sono pubblicate. Vedi foglio 08"),
    ("", None, None, None),
    ("FATTORI ANAMNESTICI (0 = assente, 1 = presente)", None, None, None),
    ("Sigarette al giorno", 20, "n/die", A1 + " / " + A2),
    ("Anni di abitudine tabagica", 30, "anni", A1 + " / " + A2),
    ("Miosteatosi documentata", 1, "0/1", A2 + " - infiltrazione adiposa del muscolo; non misurabile con l'antropometria di superficie"),
    ("Sospetta apnea ostruttiva del sonno mai indagata", 1, "0/1", "Russamento e pause respiratorie riferite dalla convivente, nessun accertamento eseguito"),
    ("Sedentarieta / assenza di esercizio strutturato", 1, "0/1", A2 + " - calcetto interrotto a 40 anni e mai ripreso, 30.000 km l'anno in auto"),
    ("", None, None, None),
    ("PARAMETRI DI CONFRONTO E SOGLIE", None, None, None),
    ("FC a riposo di riferimento ottimale", 65, "bpm", "Punto medio della fascia 60-69 bpm, minimo di mortalita (PMID 11337213)"),
    ("Soglia FC a riposo ad alto rischio", 80, "bpm", "Soglia oltre la quale il rischio cambia scala (PMID 26598376, PMID 28067310)"),
    ("Limite superiore FC del referto", 76, "bpm", "Estremo alto dell'intervallo del referto; base del calcolo dei battiti in eccesso"),
    ("Soglia vita per sindrome metabolica (maschio europeo)", 101, "cm", A2 + " - criterio antropometrico; il soggetto ne misura 97 e la casella resta vuota"),
]
r = 5
IR = {}
for name, val, unit, note in INPUTS:
    if name == "":
        r += 1
        continue
    if val is None:
        c = ws.cell(row=r, column=1, value=name)
        c.font = Font(name=FONT, size=10, bold=True, color="1F3864")
        for cc in range(1, 5):
            ws.cell(row=r, column=cc).fill = SUB_FILL
            ws.cell(row=r, column=cc).border = BORDER
        r += 1
        continue
    ws.cell(row=r, column=1, value=name).font = Font(name=FONT, size=10)
    vc = ws.cell(row=r, column=2, value=val)
    vc.font = Font(name=FONT, size=10, bold=True, color=BLUE)
    vc.fill = PatternFill("solid", fgColor=YEL)
    vc.alignment = Alignment(horizontal="center")
    uc = ws.cell(row=r, column=3, value=unit)
    uc.font = Font(name=FONT, size=10)
    uc.alignment = Alignment(horizontal="center")
    nc = ws.cell(row=r, column=4, value=note)
    nc.font = Font(name=FONT, size=9, italic=True, color="595959")
    nc.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = rowh(note, 88, 15, 44)
    for cc in range(1, 5):
        ws.cell(row=r, column=cc).border = BORDER
    IR[name] = r
    r += 1

I = lambda k: f"'01 Input'!$B${IR[k]}"
I_AGE, I_SEX = I("Eta"), I("Sesso")
I_W, I_H = I("Peso"), I("Altezza")
I_WC, I_HIP = I("Circonferenza vita"), I("Circonferenza fianchi")
I_CALF, I_WRIST = I("Circonferenza polpaccio"), I("Circonferenza polso")
I_W30 = I("Peso all'eta di 30 anni")
I_BSA = I("Superficie corporea (BSA)")
I_SBP, I_DBP = I("Pressione sistolica (SBP)"), I("Pressione diastolica (DBP)")
I_HR, I_SV = I("Frequenza cardiaca a riposo (FC)"), I("Gittata sistolica (SV)")
I_HRMAX = I("Frequenza cardiaca massima da referto")
I_CVP = I("Pressione venosa centrale assunta (CVP)")
I_HDL, I_TG = I("Colesterolo HDL"), I("Trigliceridi")
I_FE, I_TRF, I_FER = I("Sideremia"), I("Transferrina"), I("Ferritina")
I_KFE = I("Fattore di aggiustamento BRINDA - sideremia")
I_KFER = I("Fattore di aggiustamento BRINDA - ferritina")
I_PCR, I_VES = I("Proteina C reattiva (PCR)"), I("VES a 1 ora")
I_HCY, I_URIC = I("Omocisteina"), I("Uricemia")
I_PLT, I_LYM = I("Piastrine"), I("Linfociti")
I_NEU, I_MON = I("Neutrofili"), I("Monociti")
I_MPV, I_RDW, I_HB = I("Volume piastrinico medio (MPV)"), I("Ampiezza di distribuzione eritrocitaria (RDW)"), I("Emoglobina")
I_ABSIZ = I("ABSI z-score da referto")
I_CIG, I_YRS = I("Sigarette al giorno"), I("Anni di abitudine tabagica")
I_MIO = I("Miosteatosi documentata")
I_OSAS = I("Sospetta apnea ostruttiva del sonno mai indagata")
I_SED = I("Sedentarieta / assenza di esercizio strutturato")
I_HROPT = I("FC a riposo di riferimento ottimale")
I_HRTHR = I("Soglia FC a riposo ad alto rischio")
I_HRREF = I("Limite superiore FC del referto")
I_WCMS = I("Soglia vita per sindrome metabolica (maschio europeo)")

ws.cell(row=IR["Pressione venosa centrale assunta (CVP)"], column=2).comment = Comment(
    "Assunzione dell'analista: le fonti non riportano la CVP.\n"
    "Con CVP = 5 mmHg la formula 80*(MAP-CVP)/CI restituisce 2696 dyn*s*cm-5*m2 = 2,7 kU*m2,\n"
    "cioe il valore stampato sul referto. Con CVP = 0 si otterrebbe 2,8.", "Analisi")
ws.cell(row=IR["Peso"], column=2).comment = Comment(
    "ATTENZIONE: 74 kg su 174 cm danno un BMI di 24,44, non i 24,09 stampati sul referto.\n"
    "Il referto usa internamente 24,09 (lo si deduce dal VAI, che torna solo con quel valore).\n"
    "Vedi foglio '08 Controlli', riga 1.", "Analisi")

# =====================================================================
# 02 INDICI
# =====================================================================
ws = wb.create_sheet("02 Indici")
setw(ws, {"A": 5, "B": 15, "C": 36, "D": 34, "E": 13, "F": 12, "G": 13, "H": 11,
          "I": 10, "J": 10, "K": 12, "L": 11, "M": 15, "N": 7, "O": 8, "P": 9,
          "Q": 26, "R": 66})
title(ws, "02 - Indici estratti dalle fonti, ricalcolati e valutati",
      "Colonna E ricalcolata dai dati grezzi del foglio 01. Colonna G riporta il valore pubblicato: la colonna H rende visibile ogni scostamento.")
hdr(ws, 4, ["#", "Gruppo", "Indice", "Formula", "Valore\ncalcolato", "Unita", "Valore\nreferto",
            "Scost.\n%", "Rif.\nmin", "Rif.\nmax", "Direzione\ncritica", "Posiz.\nrange",
            "Stato", "Peso", "Punt.\n0-3", "Contrib.", "Fonte", "Note e interpretazione"])

R0 = 5
BMI = f"E{R0+16}"          # riga del BMI (art.2)
MAP = f"E{R0+2}"
SEXM = f'IF({I_SEX}="M",1,0)'

IND = [
 # ---------------- ARTICOLO 1 : CARDIOLOGICI ----------------
 (A1, "Pressori", "Pressione sistolica", "misurata", f"={I_SBP}", "mmHg", 128, 100, 129, "Alto", 1, "0",
  "Dentro l'intervallo ma a 1 mmHg dal limite superiore."),
 (A1, "Pressori", "Pressione diastolica", "misurata", f"={I_DBP}", "mmHg", 84, 60, 84, "Alto", 1, "0",
  "Esattamente sul limite superiore. La fonte segnala che la soglia contera quando entrera in uno score."),
 (A1, "Pressori", "Pressione arteriosa media (MAP)", "DBP + 0,4 x PP", f"=E{R0+1}+0.4*E{R0+3}", "mmHg", 101, 70, 105, "Alto", 1, "0",
  "La MAP classica (SBP+2xDBP)/3 darebbe 98,7 e non 101: il valore pubblicato si riproduce solo con DBP + 0,4 x PP. La scelta si propaga sul Modified Shock Index, che ha la MAP al denominatore."),
 (A1, "Pressori", "Pressione di pulsazione (PP)", "SBP - DBP", f"=E{R0}-E{R0+1}", "mmHg", 44, 30, 55, "Alto", 2, "0",
  "Marcatore di elasticita dei grossi vasi e di invecchiamento vascolare (PMID 33793325). Dentro il range: da sorvegliare, non da segnalare."),
 (A1, "Dinamici", "Frequenza cardiaca a riposo", "misurata", f"={I_HR}", "bpm", 82, 72, 76, "Alto", 3, "0",
  "Unico indice cardiologico fuori range, liquidato come 'tachicardia lieve'. E anche quello con il supporto epidemiologico piu solido, ed e il motore aritmetico di meta degli altri indici della sezione."),
 (A1, "Dinamici", "Frequenza cardiaca massima stimata", "modello logistico sesso-specifico", f"={I_HRMAX}", "bpm", 175, 158, 192, "Informativo", 0, "0",
  "Stima anagrafica, non misura: non descrive lo stato del soggetto, quindi peso 0. L'intervallo 158-192 e il 90-110% del predetto. Confronto con 220-eta (170) e Tanaka (173) nel foglio 03."),
 (A1, "Dinamici", "Frequenza cardiaca di riserva", "FC max - FC riposo", f"=E{R0+5}-E{R0+4}", "bpm", 93, 70, 120, "Basso", 1, "0",
  "Formalmente normale, ma alta perche la FC massima e alta per eta, non perche il basale sia buono. Con basale 72 salirebbe a 103 senza alcun miglioramento reale. HR 2,8 per mortalita CV in chi non impegna l'80% della riserva (PMID 17446799)."),
 (A1, "Dinamici", "Modified Shock Index (MSI)", "FC riposo / MAP", f"=E{R0+4}/E{R0+2}", "-", 0.81, 0.70, 0.99, "Alto", 1, "0.00",
  "Va calcolato sulla pressione MEDIA, non sulla sistolica: con la sistolica verrebbe 0,64, errore diffuso nei calcolatori online."),
 (A1, "Dinamici", "Prodotto cardiovascolare (RPP)", "FC riposo x SBP", f"=E{R0+4}*E{R0}", "bpm*mmHg", 10496, 10000, 14999, "Alto", 2, "#,##0",
  "Surrogato non invasivo del consumo miocardico di ossigeno (PMID 38453019). Il soggetto e entrato nella fascia per 496 punti, spinto dalla frequenza. Tre mesi di esercizio supervisionato lo abbassano del 19% (PMID 15793048)."),
 (A1, "Pompa", "Gittata sistolica (SV)", "misurata", f"={I_SV}", "mL", 66, 60, 100, "Basso", 2, "0",
  "Nel 15% piu basso dell'intervallo: e la pompa poco capiente che obbliga il cuore a compensare con la frequenza. E l'indice che spiega perche gli 82 battiti esistono."),
 (A1, "Pompa", "Portata cardiaca (CO)", "SV x FC / 1000", f"=E{R0+9}*E{R0+4}/1000", "L/min", 5.4, 3.5, 6.9, "Basso", 1, "0.0",
  "Il risultato finale e adeguato, ed e per questo che il referto scrive 'normale'. Lo stesso valore si ottiene con SV 80 mL e 68 bpm: identico output, meccanismo diverso."),
 (A1, "Pompa", "Cardiac Index (CI)", "CO / BSA", f"=E{R0+10}/{I_BSA}", "L/min/m2", 2.8, 2.5, 3.9, "Basso", 1, "0.00",
  "Portata indicizzata per superficie corporea. Nel terzo inferiore dell'intervallo."),
 (A1, "Pompa", "Left Cardiac Work Index (LCWI)", "CI x MAP x 2,222", f"=E{R0+11}*E{R0+2}*2.222", "mW/m2", 627, 400, 900, "Bilaterale", 1, "0",
  "La costante 2,222 converte L/min x mmHg in milliwatt (1 L/min = 1,667e-5 m3/s; 1 mmHg = 133,322 Pa). Ricalcolando con il CI non arrotondato si ottiene 639 contro i 627 pubblicati: differenza di solo arrotondamento."),
 (A1, "Pompa", "Indice di rigidita arteriosa", "PP / MAP", f"=E{R0+3}/E{R0+2}", "-", 0.44, 0.30, 0.55, "Alto", 2, "0.00",
  "Rapporto fra componente pulsatile e continua. Ricostruito per riproduzione numerica: PP/MAP da 0,436, mentre PP/SBP darebbe 0,34."),
 (A1, "Pompa", "Resistenze vascolari sistemiche (SVRI)", "80 x (MAP - CVP) / CI / 1000", f"=80*(E{R0+2}-{I_CVP})/E{R0+11}/1000", "kU*m2", 2.7, 1.8, 2.8, "Alto", 2, "0.00",
  "A 0,1 dal limite superiore: le arterie oppongono resistenza vicina al massimo tollerato e il cuore risponde con la frequenza, non aumentando il volume di eiezione. CVP non pubblicata, assunta a 5 mmHg."),
 (A1, "Antropom.", "Superficie corporea (BSA)", "input (formula non pubblicata)", f"={I_BSA}", "m2", 1.9, 1.5, 2.3, "Informativo", 0, "0.00",
  "Inclusa come sedicesima riga: la fonte dichiara 16 indici cardiologici ma ne tabula 15. E il denominatore di CI, LCWI e SVRI. Peso 0."),

 # ---------------- ARTICOLO 2 : ANTROPOMETRICI ----------------
 (A2, "Base", "Indice di massa corporea (BMI)", "peso / altezza^2", f"={I_W}/({I_H}/100)^2", "kg/m2", 24.09, 18.5, 24.9, "Informativo", 0, "0.00",
  "ATTENZIONE: 74 kg su 174 cm danno 24,44, non i 24,09 pubblicati. Il referto usa internamente 24,09 (il VAI torna solo con quel valore). Peso 0 non perche sia irrilevante, ma perche e proprio l'indice che qui assolve un corpo in disordine. Vedi foglio 08."),
 (A2, "Base", "Rapporto altezza / polso", "altezza / circonf. polso", f"={I_H}/{I_WRIST}", "-", 9.7, 9.6, 10.4, "Informativo", 0, "0.00",
  "Morfotipo normolineo. Al polso non c'e ne grasso apprezzabile ne ventre muscolare: e l'unica misura stabile per tutta la vita adulta, e quindi il riferimento fisso contro cui si giudicano gli indici che invece si muovono."),
 (A2, "Base", "Peso ideale secondo Lorenz", "H - 100 - (H - 150)/4", f"={I_H}-100-({I_H}-150)/4", "kg", 68.0, 60, 80, "Informativo", 0, "0.0",
  "Lavora sulla sola statura e ignora eta e corporatura. Estremo basso della forbice del peso desiderabile."),
 (A2, "Base", "Peso ideale secondo Creff", "(H - 100 + eta/10) x 0,9", f"=({I_H}-100+{I_AGE}/10)*0.9", "kg", 71.1, 60, 80, "Informativo", 0, "0.0",
  "Corregge per eta e morfotipo. Con Lorenz definisce la forbice 68-71 kg: il soggetto a 74 kg sfora di meno di tre chili, margine che nessun clinico definirebbe allarmante. Il problema non e quanto pesa ma dove."),
 (A2, "Forma", "Rapporto vita / altezza (WHtR)", "vita / altezza", f"={I_WC}/{I_H}", "-", 0.56, 0.30, 0.52, "Alto", 3, "0.00",
  "La regola aurea vuole la vita sotto la meta dell'altezza. Meta-analisi di 31 studi: discrimina il rischio cardiometabolico meglio del BMI (PMID 22106927). E il primo indice da guardare, per velocita e robustezza."),
 (A2, "Forma", "Rapporto vita / fianchi (WHR)", "vita / fianchi", f"={I_WC}/{I_HIP}", "-", 0.97, 0.70, 0.89, "Alto", 2, "0.00",
  "Forma a mela: distribuzione addominale del grasso, la peggiore dal punto di vista metabolico."),
 (A2, "Forma", "Conicity Index", "(vita/100) / (0,109 x RADQ(peso/altezza_m))", f"=({I_WC}/100)/(0.109*SQRT({I_W}/({I_H}/100)))", "-", 1.36, 1.00, 1.25, "Alto", 2, "0.00",
  "Descrive quanto il tronco si allontana dal cilindro per avvicinarsi al doppio cono. Normalizza la vita per peso e statura, quindi l'addome sporge al netto della corporatura."),
 (A2, "Adiposita", "Percentuale di grasso corporeo", "Deurenberg: 1,2xBMI + 0,23xeta - 10,8xsesso - 5,4", f"=1.2*{BMI}+0.23*{I_AGE}-10.8*{SEXM}-5.4", "%", 23, 8, 25, "Alto", 1, "0.0",
  "Formula non dichiarata dalla fonte: Deurenberg e la piu probabile ma restituisce 24,6 contro i 23 pubblicati. Stima derivata dal BMI, quindi eredita il difetto del BMI di non distinguere i tessuti."),
 (A2, "Adiposita", "Body Adiposity Index (BAI)", "fianchi / altezza_m^1,5 - 18", f"={I_HIP}/(({I_H}/100)^1.5)-18", "%", 25, 8, 23, "Alto", 2, "0.0",
  "Ricava l'adiposita da fianchi e altezza senza mai usare il peso. Quando due indici di adiposita divergono sulla stessa persona vince quello che non usa il peso, perche e il peso a mascherare la ricomposizione corporea."),
 (A2, "Adiposita", "Body Roundness Index (BRI)", "364,2 - 365,5 x RADQ(1 - ((vita/2pi)/(0,5xH))^2)", f"=364.2-365.5*SQRT(1-((({I_WC}/100)/(2*PI()))/(0.5*({I_H}/100)))^2)", "-", 4.50, 1.0, 4.71, "Alto", 1, "0.00",
  "Uno dei due indici assolti del gruppo, e va registrato con onestà. Dentro il riferimento, ma nell'ultimo 10% dell'intervallo."),
 (A2, "Adiposita", "Abdominal Volume Index (AVI)", "(2 x vita^2 + 0,7 x (vita-fianchi)^2) / 1000", f"=(2*{I_WC}^2+0.7*({I_WC}-{I_HIP})^2)/1000", "-", 11.89, 0, 24.5, "Alto", 1, "0.00",
  "NON RIPRODUCIBILE: la formula standard restituisce 18,82 contro gli 11,89 pubblicati. L'esito resta 'dentro il riferimento' con entrambi i valori, quindi la conclusione non cambia, ma il numero non torna. Vedi foglio 08."),
 (A2, "Ibridi", "Lipid Accumulation Product (LAP)", "(vita - 65) x trigliceridi in mmol/L", f"=({I_WC}-65)*({I_TG}/88.57)", "-", 53.5, 0, 26.7, "Alto", 3, "0.0",
  "Il doppio esatto del cut-off. Riconosce il rischio cardiovascolare meglio del BMI (PMID 16150143). Nasce dall'incrocio fra una vita di 97 cm e trigliceridi che presi da soli nessun laboratorio segnalerebbe."),
 (A2, "Ibridi", "Cardiometabolic Index (CMI)", "(vita/altezza) x (TG/HDL) in mmol/L", f"=({I_WC}/{I_H})*(({I_TG}/88.57)/({I_HDL}/38.67))", "-", 0.88, 0, 0.39, "Alto", 2, "0.00",
  "Piu del doppio del cut-off. Buon discriminatore di diabete e disglicemia (PMID 25199852). Va calcolato in mmol/L: usando i mg/dL verrebbe 2,01, un numero che non ha piu alcun rapporto con la soglia."),
 (A2, "Ibridi", "Visceral Adiposity Index (VAI)", "(vita/(39,68+1,88xBMI)) x (TG/1,03) x (1,31/HDL), mmol/L", f"=({I_WC}/(39.68+1.88*{BMI}))*(({I_TG}/88.57)/1.03)*(1.31/({I_HDL}/38.67))", "-", 2.29, 0, 1.92, "Alto", 3, "0.00",
  "Marcatore della FUNZIONE del tessuto adiposo viscerale, non della sua quantita (PMID 20067971). Il referto commenta: disfunzione moderata del tessuto adiposo, resistenza insulinica presente. E l'indice che rivela il BMI usato internamente dal referto."),
 (A2, "Mortalita", "A Body Shape Index (ABSI) - z-score", "z da tabelle NHANES per eta e sesso", f"={I_ABSIZ}", "z", 1.52, -0.272, 0.228, "Alto", 3, "0.00",
  "L'unico indice il cui commento non ammette diplomazia: rischio di mortalita molto alto. Costruito su 14.105 adulti NHANES per isolare la forma corporea depurata da peso e BMI, con aumento quasi esponenziale della mortalita (PMID 22815707). E l'indice che un normopeso non si aspetta mai di trovare alterato. Il valore grezzo di ABSI e ricalcolabile (0,0873), lo z-score no: le tabelle di riferimento non sono pubblicate."),
 (A2, "Muscolo", "Circonferenza polpaccio corretta per BMI", "CC - correzione categoriale EWGSOP2", f"={I_CALF}-IF({BMI}<18.5,-4,IF({BMI}<25,0,IF({BMI}<30,3,7)))", "cm", 33.5, 34.0, 45.0, "Basso", 3, "0.0",
  "Surrogato validato della massa muscolare appendicolare (PMID 30312372). La correzione applicata dal referto non e pubblicata: con quella categoriale EWGSOP2 il BMI del soggetto non comporta alcuna correzione e il valore resterebbe 34,0 cm, cioe ESATTAMENTE sulla soglia invece che sotto. La diagnosi di sarcopenia moderata dipende interamente da quei 5 millimetri non documentati."),

 # ---------------- ARTICOLO 3 : PROFILO MARZIALE ----------------
 (A3, "Ferro circolante", "Sideremia", "misurata", f"={I_FE}", "uM", 13.6, 13, 27, "Basso", 2, "0.0",
  "Dentro l'intervallo per sei decimi di micromole: e il valore piu basso possibile senza essere segnalato. Ha un ritmo circadiano marcato, quindi un prelievo pomeridiano abbasserebbe la saturazione di diversi punti senza che le riserve siano cambiate."),
 (A3, "Ferro circolante", "Sideremia aggiustata per infiammazione", "sideremia x fattore BRINDA", f"={I_FE}*{I_KFE}", "uM", 17.1, 15, 23, "Basso", 1, "0.0",
  "La correzione per l'infiammazione alza il ferro circolante e abbassa la ferritina, cioe spinge le due grandezze in direzioni opposte. Il fattore non e pubblicato: e stato dedotto dal rapporto fra valore aggiustato e valore grezzo."),
 (A3, "Trasporto", "Transferrina", "misurata", f"={I_TRF}", "uM", 25.1, 25.1, 50.3, "Alto", 1, "0.0",
  "Il camion che trasporta il ferro. Il fegato ne produce di piu proprio quando il ferro scarseggia, quindi un valore alto e un segnale di carenza, non di abbondanza."),
 (A3, "Trasporto", "TIBC (capacita totale legante)", "2 x transferrina", f"=2*E{R0+34}", "uM", 78.4, 45, 76, "Alto", 1, "0.0",
  "Ogni molecola di transferrina lega due atomi di ferro: da qui il fattore 2, che riproduce esattamente il valore pubblicato. Oltre la finestra ottimale significa che circolano molti posti liberi sui camion."),
 (A3, "Trasporto", "UIBC (capacita legante insatura)", "TIBC - sideremia", f"=E{R0+35}-E{R0+32}", "uM", 64.8, 22, 61, "Alto", 1, "0.0",
  "I posti liberi effettivi. Conferma dall'altro lato quello che dice la saturazione: la capacita di trasporto c'e, il carico no."),
 (A3, "Disponibilita", "Saturazione della transferrina", "sideremia / TIBC x 100", f"=E{R0+32}/E{R0+35}*100", "%", 17.3, 20, 48, "Basso", 3, "0.0",
  "L'UNICO valore dell'intero pannello sotto il riferimento anche senza aggiustamenti, ed e considerato il marcatore piu informativo dello stato marziale reale perche mette insieme ferro circolante e proteina di trasporto (PMID 27346617, PMID 36380788). I camion viaggiano pieni per un sesto quando dovrebbero stare almeno a un quinto."),
 (A3, "Disponibilita", "Saturazione aggiustata", "sideremia aggiustata / TIBC x 100", f"=E{R0+33}/E{R0+35}*100", "%", 21.8, 24, 35, "Basso", 2, "0.0",
  "Anche dopo la correzione che alza il ferro circolante, la saturazione resta sotto la finestra ottimale. La carenza funzionale non e un artefatto dell'infiammazione."),
 (A3, "Depositi", "Ferritina", "misurata", f"={I_FER}", "pM", 160.9, 54, 854, "Informativo", 0, "0.0",
  "Comoda nel mezzo di un intervallo larghissimo, ed e la riga su cui quasi tutti si fermano. Peso 0 come il BMI dell'articolo 2, e per la stessa ragione: e una proteina di fase acuta, quindi in un fumatore da trent'anni con flogosi cronica misura l'infiammazione piu delle riserve. Correggendo per infiammazione con il metodo BRINDA la prevalenza di carenza sale dal 46,3% al 61,5% (PMID 35623855)."),
 (A3, "Depositi", "Ferritina aggiustata per infiammazione", "ferritina x fattore BRINDA", f"={I_FER}*{I_KFER}", "pM", 120.7, 283.1, 854, "Basso", 3, "0.0",
  "Meno della meta del valore ottimale. E la riga che nessuno guarda e che ribalta la lettura della riga precedente: il commento del laboratorio sulla stessa fascia dice depositi midollari in esaurimento."),
 (A3, "Depositi", "Depositi di ferro", "ferritina x 4,5", f"={I_FER}*4.5", "nM", 724, 301.5, 706.5, "Bilaterale", 1, "0.0",
  "Il fattore 4,5 deriva dai circa 4500 atomi di ferro che ogni molecola di ferritina puo contenere, con la conversione da pM a nM. Sopra la finestra ottimale: ferro presente ma sequestrato, non ferro in eccesso. Calcolandolo sulla ferritina AGGIUSTATA darebbe 543 nM, cioe dentro la finestra: la contraddizione fra le due righe e il cuore del referto non discriminante."),
 (A3, "Indici derivati", "Transferrina / log(ferritina)", "transferrina g/L / log10(ferritina ug/L)", f"=({I_TRF}*{K_TRF_GL})/LOG10({I_FER}*{K_FER_UGL})", "-", 1.63, 0, 1.70, "Alto", 2, "0.00",
  "Indice combinato per la carenza marziale in presenza di infiammazione. Il valore dipende dai pesi molecolari usati per la conversione in unita convenzionali, che la fonte non dichiara: al variare di quelli plausibili il risultato oscilla fra 1,59 e 1,70, e l'1,63 pubblicato cade dentro questa banda. Comunque a ridosso del cut-off."),

 # ---------------- ARTICOLO 4 : INDICI INFIAMMATORI ----------------
 (A4, "Classici", "Proteina C reattiva (PCR)", "misurata", f"={I_PCR}", "mg/dL", 0.3, 0, 0.5, "Informativo", 0, "0.0",
  "Peso 0 come il BMI e come la ferritina, e per la stessa ragione: e un esame giusto per la domanda sbagliata. Sale molto e in fretta nell'infiammazione acuta, quindi non intercetta l'infiammazione cronica di basso grado, che non e un incendio ma un fornello lasciato acceso per vent'anni."),
 (A4, "Classici", "VES a 1 ora", "misurata", f"={I_VES}", "mm", 14, 1, 25, "Informativo", 0, "0",
  "Stesso limite della PCR con in piu una lentezza propria, perche misura una proprieta reologica del sangue che cambia nell'arco di giorni. Peso 0."),
 (A4, "Rapporti cellulari", "Neutrofili / Linfociti (NLR)", "neutrofili / linfociti", f"={I_NEU}/{I_LYM}", "-", 3.4, 0.73, 3.33, "Alto", 2, "0.00",
  "L'infiammazione cronica sposta due popolazioni in direzioni opposte: alza l'immunita innata e abbassa i linfociti. Il rapporto cattura lo spostamento anche quando entrambi i valori assoluti restano dentro i loro intervalli, che e esattamente il caso qui (neutrofili 69% con limite 70, linfociti 20,5% con limite 20)."),
 (A4, "Rapporti cellulari", "Monociti / Linfociti (MLR)", "monociti / linfociti", f"={I_MON}/{I_LYM}", "-", 0.38, 0.12, 0.38, "Alto", 2, "0.00",
  "Esattamente sul limite superiore, come la pressione diastolica dell'articolo 1. L'articolo lo conta fra i dieci valori fuori riferimento; il modello lo classifica come borderline perche non lo supera. E l'unica differenza fra il conteggio della fonte e quello del foglio 06."),
 (A4, "Rapporti cellulari", "Piastrine / Linfociti (PLR)", "piastrine / linfociti", f"={I_PLT}/{I_LYM}", "-", 158, 63, 209, "Alto", 1, "0",
  "Dentro l'intervallo. E uno dei sei indici che non segnalano nulla, e va registrato."),
 (A4, "Indici compositi", "Systemic Inflammation Index (SII)", "piastrine x neutrofili / linfociti", f"={I_PLT}*{I_NEU}/{I_LYM}", "-", 917, 131, 901, "Alto", 2, "0",
  "Combina le tre popolazioni in un solo numero. Appena oltre il limite superiore. Su 42.875 adulti seguiti per vent'anni SII e SIRI elevati risultano associati a mortalita totale e cardiovascolare (PMID 36769776)."),
 (A4, "Indici compositi", "Systemic Inflammation Response Index (SIRI)", "neutrofili x monociti / linfociti", f"={I_NEU}*{I_MON}/{I_LYM}", "-", 2.23, 0, 0.68, "Alto", 3, "0.00",
  "Tre volte e mezzo la soglia. Insieme all'AISI e lo scostamento piu largo di tutto il workbook, in un uomo che sull'emocromo aveva un solo parametro fuori su ventisette."),
 (A4, "Indici compositi", "Aggregate Index of Systemic Inflammation (AISI)", "neutrofili x monociti x piastrine / linfociti", f"={I_NEU}*{I_MON}*{I_PLT}/{I_LYM}", "-", 606.56, 0, 147.16, "Alto", 3, "0.00",
  "Quattro volte la soglia. Su 23.765 ipertesi il quartile piu alto mostrava un rischio di mortalita cardiovascolare quasi doppio rispetto al piu basso, con hazard ratio 1,91 (PMID 37265570). Gli indici derivati dall'emocromo sono stati associati anche alla sarcopenia (PMID 38755603), che chiude il cerchio con l'articolo 2."),
 (A4, "Rapporti misti", "Monociti / HDL", "monociti / HDL x 1000", f"={I_MON}/{I_HDL}*1000", "per mille", 16.1, 0, 6, "Alto", 3, "0.0",
  "Quasi il triplo del limite. Le HDL non trasportano solo colesterolo: hanno una funzione antinfiammatoria diretta e ostacolano il reclutamento dei monociti nella parete arteriosa. Il rapporto mette in relazione chi accende l'infiammazione e chi dovrebbe spegnerla. Correlato a PCR, conta leucocitaria e ipertensione resistente (PMID 34109496). L'HDL a 41 dell'articolo 2 qui smette di essere un dettaglio."),
 (A4, "Rapporti misti", "Emoglobina / RDW", "emoglobina / RDW", f"={I_HB}/{I_RDW}", "-", 0.96, 1.0, 2.0, "Basso", 2, "0.00",
  "Sotto il valore minimo. Mette insieme quanta emoglobina c'e e quanto sono disomogenei i globuli rossi: cala quando il midollo produce eritrociti di taglia irregolare, che e il quadro dell'eritropoiesi ferro-ristretta dell'articolo 3."),
 (A4, "Rapporti misti", "MPV / Linfociti", "volume piastrinico medio / linfociti", f"={I_MPV}/{I_LYM}", "-", 6.84, 0, 5.55, "Alto", 1, "0.00",
  "Il volume piastrinico medio sale quando il midollo rilascia piastrine piu giovani e reattive, ed e un marcatore di attivazione piastrinica nell'infiammazione (PMID 31148950)."),
 (A4, "Rapporti misti", "MPV / Piastrine", "MPV / piastrine x 100", f"={I_MPV}/{I_PLT}*100", "-", 4.3, 0, 4.0, "Alto", 1, "0.00",
  "Appena oltre la soglia. Descrive lo stesso fenomeno del precedente rapportato alla massa piastrinica totale."),
 (A4, "Rapporti misti", "RDW / Piastrine", "RDW / piastrine", f"={I_RDW}/{I_PLT}", "-", 0.05, 0, 0.07, "Alto", 1, "0.000",
  "Dentro la soglia. E pubblicato con due sole cifre decimali, il che lascia molta indeterminazione: e da questa riga che dipende la ricostruzione di RDW ed emoglobina nel foglio 01."),
 (A4, "Metabolici", "Omocisteina", "misurata", f"={I_HCY}", "uM", 14.6, 3, 15, "Alto", 2, "0.0",
  "Dentro l'intervallo largo per quattro decimi, ma il doppio del limite superiore della finestra ottimale (5,0-7,2) indicata dalla stessa tabella. E il divario piu ampio fra riferimento tollerato e riferimento desiderabile di tutto il workbook."),
 (A4, "Metabolici", "Uricemia", "misurata", f"={I_URIC}", "mg/dL", 6.7, 3.5, 7.2, "Alto", 2, "0.0",
  "Dentro l'intervallo largo, ben oltre la finestra ottimale (2,5-4,0). L'acido urico e a sua volta un promotore di infiammazione e di disfunzione endoteliale."),
 (A4, "Metabolici", "Uricemia / HDL", "uricemia / HDL x 100", f"={I_URIC}/{I_HDL}*100", "-", 16.3, 0, 12.2, "Alto", 2, "0.0",
  "Un terzo sopra il cut-off. Come il rapporto monociti/HDL, deve il suo valore tanto al numeratore alto quanto al denominatore basso: l'HDL a 41 compare come denominatore in due dei tre indici piu alterati di questa sezione."),
]

r = R0
for i, (fonte, grp, nome, ftxt, fexc, unit, ref, mn, mx, dz, peso, nfmt, nota) in enumerate(IND, 1):
    ws.cell(row=r, column=1, value=i)
    ws.cell(row=r, column=2, value=grp)
    ws.cell(row=r, column=3, value=nome).font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=4, value=ftxt).font = Font(name=FONT, size=8, italic=True)
    ce = ws.cell(row=r, column=5, value=fexc)
    ce.font = Font(name=FONT, size=10, bold=True, color=BLACK)
    ce.number_format = nfmt
    ws.cell(row=r, column=6, value=unit).font = Font(name=FONT, size=9)
    cg = ws.cell(row=r, column=7, value=ref)
    cg.font = Font(name=FONT, size=10, color=BLUE)
    cg.number_format = nfmt
    ws.cell(row=r, column=8, value=f'=IF(G{r}=0,"",(E{r}-G{r})/ABS(G{r}))').number_format = "0.0%"
    for col, v in ((9, mn), (10, mx)):
        c = ws.cell(row=r, column=col, value=v)
        c.number_format = nfmt
        c.font = Font(name=FONT, size=10, color=BLUE)
    ws.cell(row=r, column=11, value=dz).font = Font(name=FONT, size=9)
    ws.cell(row=r, column=12, value=f'=IF(J{r}=I{r},"",(E{r}-I{r})/(J{r}-I{r}))').number_format = "0%"
    cm = ws.cell(row=r, column=13, value=(
        f'=IF(E{r}>J{r},"Fuori (alto)",IF(E{r}<I{r},"Fuori (basso)",'
        f'IF(AND(K{r}<>"Informativo",OR(AND(OR(K{r}="Alto",K{r}="Bilaterale"),L{r}>0.8),'
        f'AND(OR(K{r}="Basso",K{r}="Bilaterale"),L{r}<0.2))),"Borderline","Normale")))'))
    cm.font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=14, value=peso).font = Font(name=FONT, size=10, color=BLUE)
    co = ws.cell(row=r, column=15, value=(
        f'=IF(N{r}=0,0,IF(AND(E{r}>J{r},OR(K{r}="Alto",K{r}="Bilaterale")),IF((E{r}-J{r})/(J{r}-I{r})>0.1,3,2),'
        f'IF(AND(E{r}<I{r},OR(K{r}="Basso",K{r}="Bilaterale")),IF((I{r}-E{r})/(J{r}-I{r})>0.1,3,2),'
        f'IF(M{r}="Borderline",1,0))))'))
    co.font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=16, value=f"=N{r}*O{r}")
    ws.cell(row=r, column=17, value=fonte).font = Font(name=FONT, size=8, color="595959")
    cn = ws.cell(row=r, column=18, value=nota)
    cn.font = Font(name=FONT, size=9)
    cn.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = rowh(nota, 82, 30, 92)
    for cc in range(1, 19):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        if cc in (1, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16):
            cell.alignment = Alignment(horizontal="center", vertical="center")
        else:
            cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=(cc in (3, 4, 18)))
    r += 1
LAST = r - 1

tr = r + 1
ws.cell(row=tr, column=3, value="TOTALE PUNTEGGIO STRUMENTALE (entrambe le fonti)").font = Font(name=FONT, size=10, bold=True)
ws.cell(row=tr, column=14, value=f"=SUM(N{R0}:N{LAST})").font = Font(name=FONT, size=10, bold=True)
ws.cell(row=tr, column=16, value=f"=SUM(P{R0}:P{LAST})").font = Font(name=FONT, size=10, bold=True)
for cc in range(3, 18):
    ws.cell(row=tr, column=cc).fill = SUB_FILL
    ws.cell(row=tr, column=cc).border = BORDER

ws.freeze_panes = "E5"
ws.auto_filter.ref = f"A4:R{LAST}"
rng = f"A{R0}:R{LAST}"
ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT($M{R0},5)="Fuori"'], fill=BAD_FILL))
ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$M{R0}="Borderline"'], fill=WARN_FILL))
ws.conditional_formatting.add(f"H{R0}:H{LAST}",
                              FormulaRule(formula=[f'ABS($H{R0})>0.02'], fill=ERR_FILL,
                                          font=Font(name=FONT, bold=True, color="C00000")))

C = lambda off: f"'02 Indici'!$E${R0+off}"
E_WHtR, E_LAP, E_CMI, E_VAI, E_CC = C(20), C(27), C(28), C(29), C(31)
E_BMI = f"'02 Indici'!${BMI}"
E_SAT, E_FERADJ, E_DEP = C(37), C(40), C(41)
E_PCR, E_SIRI, E_AISI, E_MHR = C(43), C(49), C(50), C(51)

# =====================================================================
# 03 FORMULE
# =====================================================================
ws = wb.create_sheet("03 Formule")
setw(ws, {"A": 5, "B": 13, "C": 32, "D": 44, "E": 26, "F": 18, "G": 66})
title(ws, "03 - Dizionario delle formule",
      "Le formule marcate 'Ricostruita' non sono scritte per esteso nelle fonti: sono state ricavate cercando l'espressione che riproduce il valore pubblicato dai dati grezzi noti.")
hdr(ws, 4, ["#", "Fonte", "Grandezza", "Formula", "Verifica sul caso", "Origine", "Nota metodologica"])

FORM = [
 (A1, "Pressione arteriosa media", "MAP = DBP + 0,4 x (SBP - DBP)", "84 + 0,4x44 = 101,6 -> 101", "Ricostruita",
  "La formula classica (SBP + 2xDBP)/3 restituisce 98,7 e NON riproduce il valore pubblicato. Il coefficiente 0,4 al posto di 1/3 tiene conto dell'accorciamento della diastole. Chi ricalcola con la classica ottiene un Modified Shock Index di 0,83 invece di 0,81."),
 (A1, "Pressione di pulsazione", "PP = SBP - DBP", "128 - 84 = 44", "Esplicita",
  "Cresce quando l'aorta perde elasticita e smette di ammortizzare l'onda sistolica."),
 (A1, "FC massima - formula popolare", "FCmax = 220 - eta", "220 - 50 = 170", "Citata come errata",
  "Sottostima la FC massima negli adulti maturi. E la formula che fa sospettare un errore a chi verifica il referto."),
 (A1, "FC massima - Tanaka", "FCmax = 208 - 0,7 x eta", "208 - 35 = 173", "PMID 11153730",
  "Meta-analisi su 351 studi e 18.712 soggetti. Migliora la stima ma resta lineare, quindi applica lo stesso decadimento a ogni eta."),
 (A1, "FC massima - modello del referto", "modello logistico sesso-specifico", "175 (da referto)", "PMID 12752560 / 11722475",
  "Tarato su cicloergometro fino a esaurimento. Decadimento non costante: 0,43 bpm/anno a 25 anni, 0,83 a 50, 1,35 a 75. Coefficienti non pubblicati."),
 (A1, "FC di riserva", "FCris = FCmax - FCriposo", "175 - 82 = 93", "Esplicita",
  "Migliora abbassando il basale, cioe senza che il cuore cambi davvero: e un artefatto da tenere presente."),
 (A1, "Modified Shock Index", "MSI = FCriposo / MAP", "82 / 101 = 0,81", "Esplicita",
  "La fonte segnala l'errore diffuso di usare la sistolica al denominatore: 82/128 darebbe 0,64."),
 (A1, "Prodotto cardiovascolare", "RPP = FCriposo x SBP", "82 x 128 = 10.496", "Esplicita / PMID 38453019",
  "Rate pressure product, surrogato del consumo miocardico di ossigeno."),
 (A1, "Portata cardiaca", "CO = SV x FC", "66 x 82 = 5.412 mL/min", "Esplicita",
  "Lo stesso output si ottiene con SV 80 mL e FC 68: la portata normale non distingue una pompa capiente lenta da una piccola veloce."),
 (A1, "Cardiac Index", "CI = CO / BSA", "5,412 / 1,9 = 2,85 -> 2,8", "Esplicita", "Portata indicizzata per superficie corporea."),
 (A1, "Left Cardiac Work Index", "LCWI = CI x MAP x 2,222", "2,8 x 101 x 2,222 = 628 (referto 627)", "Ricostruita",
  "La costante converte L/min x mmHg in milliwatt: 1 L/min = 1,6667e-5 m3/s e 1 mmHg = 133,322 Pa, prodotto 2,222 mW. Con il CI non arrotondato si ottiene 639, circa il 2% sopra il pubblicato."),
 (A1, "Indice di rigidita arteriosa", "ASI = PP / MAP", "44 / 101 = 0,436 -> 0,44", "Ricostruita",
  "Il rapporto alternativo PP/SBP darebbe 0,34 e non riproduce il referto."),
 (A1, "Resistenze vascolari sistemiche", "SVRI = 80 x (MAP - CVP) / CI", "80 x (101-5) / 2,85 = 2.696 -> 2,7 kU*m2", "Ricostruita",
  "Formula standard in dyn*s*cm-5*m2, espressa dal referto in kilo-unita. CVP non pubblicata, assunta a 5 mmHg: con CVP 0 si otterrebbe 2,8."),
 (A1, "Zone di allenamento", "inizio zona = INT(% x FCmax)", "INT(0,5x175)=87; INT(0,6x175)=105", "Ricostruita",
  "Le soglie del referto si riproducono con il troncamento all'intero, non con l'arrotondamento: 0,5x175 = 87,5 diventa 87 e non 88."),
 (A1, "Battiti in eccesso all'anno", "(FC - limite sup.) x 60 x 24 x 365", "6 x 525.600 = 3.153.600", "Esplicita",
  "Traduce lo scostamento istantaneo in carico cumulativo annuo."),
 (A1, "Rischio relativo da FC a riposo", "RR = 1,09^((FC - FCrif)/10)", "1,09^((82-65)/10) = 1,16", "PMID 26598376",
  "Estrapolazione log-lineare dell'incremento per ogni 10 bpm. Stima di gradiente, non rischio assoluto individuale."),
 (A2, "Indice di massa corporea", "BMI = peso / altezza_m^2", "74 / 1,74^2 = 24,44 (referto 24,09)", "Esplicita",
  "SCOSTAMENTO: il valore pubblicato non e ottenibile dal peso e dall'altezza dichiarati. Servirebbero 72,9 kg a 174 cm, oppure 175,3 cm a 74 kg. Il BMI entra in VAI, ABSI e nella correzione del polpaccio."),
 (A2, "Morfotipo", "altezza / circonferenza polso", "174 / 18 = 9,67 -> 9,7", "Esplicita",
  "Normolineo se compreso fra 9,6 e 10,4. Al polso non c'e ne grasso apprezzabile ne ventre muscolare: e l'unica misura che non cambia con dieta o allenamento."),
 (A2, "Peso ideale - Lorenz", "P = H - 100 - (H - 150)/4", "174 - 100 - 6 = 68,0", "Ricostruita",
  "Riproduce esattamente il valore pubblicato. Usa la sola statura e ignora eta e corporatura."),
 (A2, "Peso ideale - Creff", "P = (H - 100 + eta/10) x 0,9", "(174-100+5) x 0,9 = 71,1", "Ricostruita",
  "Riproduce esattamente il valore pubblicato. Corregge per eta e morfotipo: il peso ottimale di un cinquantenne normolineo non e quello di un venticinquenne longilineo."),
 (A2, "Rapporto vita / altezza", "WHtR = vita / altezza", "97 / 174 = 0,557 -> 0,56", "Esplicita / PMID 22106927",
  "Meta-analisi di 31 studi: discrimina il rischio cardiometabolico meglio del BMI. Se supera 0,50 vale la pena calcolare tutta la batteria."),
 (A2, "Rapporto vita / fianchi", "WHR = vita / fianchi", "97 / 100 = 0,97", "Esplicita", "Cut-off 0,89 per il maschio. Forma a mela."),
 (A2, "Conicity Index", "CI = (vita_m) / (0,109 x RADQ(peso/altezza_m))", "0,97 / (0,109 x 6,522) = 1,3646 -> 1,36", "Ricostruita",
  "Riproduce esattamente il valore pubblicato. Misura quanto il tronco si allontana dal cilindro per avvicinarsi al doppio cono, al netto di peso e statura."),
 (A2, "Percentuale di grasso corporeo", "Deurenberg: 1,20xBMI + 0,23xeta - 10,8xsesso - 5,4", "24,6 con BMI ricalcolato; 24,2 con BMI 24,09 (referto 23)", "Non riproducibile",
  "La fonte non dichiara la formula. Deurenberg e la candidata piu probabile ma non restituisce il valore pubblicato con nessuno dei due BMI. Scostamento di circa 1,2 punti percentuali."),
 (A2, "Body Adiposity Index", "BAI = fianchi / altezza_m^1,5 - 18", "100 / 2,2952 - 18 = 25,57 -> 25", "Ricostruita",
  "Riproduce il valore pubblicato. Non usa mai il peso: e per questo che quando diverge dalla stima basata sul BMI, la fonte da ragione al BAI."),
 (A2, "Body Roundness Index", "BRI = 364,2 - 365,5 x RADQ(1 - ((vita/2pi)/(0,5xH))^2)", "364,2 - 359,70 = 4,5005 -> 4,50", "Ricostruita",
  "Riproduce esattamente il valore pubblicato, alla seconda cifra decimale."),
 (A2, "Abdominal Volume Index", "AVI = (2 x vita^2 + 0,7 x (vita-fianchi)^2) / 1000", "18,824 (referto 11,89)", "Non riproducibile",
  "SCOSTAMENTO del 58%. La formula standard di Guerrero-Romero e Rodriguez-Moran non restituisce il valore pubblicato. L'esito qualitativo non cambia (dentro il riferimento con entrambi i valori), ma il numero non torna."),
 (A2, "Lipid Accumulation Product", "LAP = (vita - 65) x TG in mmol/L", "32 x 1,671 = 53,47 -> 53,5", "Ricostruita / PMID 16150143",
  "Riproduce il valore pubblicato. I trigliceridi vanno convertiti in mmol/L dividendo i mg/dL per 88,57."),
 (A2, "Cardiometabolic Index", "CMI = (vita/altezza) x (TG/HDL) in mmol/L", "0,5575 x 1,576 = 0,879 -> 0,88", "Ricostruita / PMID 25199852",
  "Riproduce il valore pubblicato SOLO in mmol/L. Con i mg/dL verrebbe 2,01, che non ha piu alcun rapporto con il cut-off di 0,39."),
 (A2, "Visceral Adiposity Index", "VAI(M) = (vita/(39,68+1,88xBMI)) x (TG/1,03) x (1,31/HDL)", "2,2883 con BMI 24,09 -> 2,29; 2,2706 con BMI 24,44", "Ricostruita / PMID 20067971",
  "Riproduce il valore pubblicato solo usando il BMI di 24,09: e la prova che il referto usa internamente quel BMI e non quello ricalcolabile da peso e altezza. TG e HDL in mmol/L."),
 (A2, "A Body Shape Index", "ABSI = vita_m / (BMI^(2/3) x RADQ(altezza_m))", "0,97 / 11,11 = 0,0873", "Parzialmente ricostruita / PMID 22815707",
  "Il valore grezzo si ricalcola. Lo z-score pubblicato (1,52) richiede le tabelle NHANES per eta e sesso, non riportate dalla fonte, quindi nel foglio 02 viene inserito a mano."),
 (A2, "Circonferenza polpaccio corretta", "CC - correzione categoriale per BMI (EWGSOP2)", "34,0 - 0 = 34,0 (referto 33,5)", "Non riproducibile",
  "La correzione categoriale non prevede alcun aggiustamento per BMI fra 18,5 e 24,9, quindi il valore resterebbe 34,0 cm, esattamente sulla soglia. Il referto applica -0,5 cm con una formula non pubblicata, e da quei 5 millimetri dipende la diagnosi di sarcopenia moderata."),
 (A2, "Volume del tessuto adiposo", "a parita di massa il grasso occupa circa il 18% di spazio in piu del muscolo", "-", "Esplicita",
  "Spiega perche la vita cresce mentre l'ago della bilancia sta fermo, e perche il paziente giura in buona fede di non essere ingrassato."),
 (A3, "TIBC", "TIBC = 2 x transferrina", "2 x 39,2 = 78,4", "Ricostruita",
  "Riproduce esattamente il valore pubblicato. Il fattore 2 e stechiometrico: ogni molecola di transferrina lega due atomi di ferro."),
 (A3, "UIBC", "UIBC = TIBC - sideremia", "78,4 - 13,6 = 64,8", "Ricostruita",
  "Riproduce esattamente il valore pubblicato. E la capacita di trasporto ancora libera."),
 (A3, "Saturazione della transferrina", "TSAT = sideremia / TIBC x 100", "13,6 / 78,4 = 17,35% -> 17,3", "Ricostruita",
  "Riproduce il valore pubblicato. E l'unico valore del pannello sotto il riferimento senza bisogno di correzioni, ed e considerato il marcatore piu informativo dello stato marziale reale."),
 (A3, "Saturazione aggiustata", "TSATadj = sideremia aggiustata / TIBC x 100", "17,1 / 78,4 = 21,81% -> 21,8", "Ricostruita",
  "Riproduce il valore pubblicato e conferma che l'aggiustamento agisce sul numeratore, non sulla capacita di trasporto."),
 (A3, "Depositi di ferro", "depositi (nM) = ferritina (pM) x 4,5", "160,9 x 4,5 = 724,05 -> 724", "Ricostruita",
  "Riproduce esattamente il valore pubblicato. Il fattore 4,5 deriva dai circa 4500 atomi di ferro per molecola di ferritina, diviso 1000 per passare da picomoli a nanomoli. Nota: il calcolo usa la ferritina GREZZA; con quella aggiustata darebbe 543 nM, cioe dentro la finestra ottimale invece che sopra."),
 (A3, "Aggiustamento per infiammazione", "valore aggiustato = valore grezzo x fattore BRINDA", "ferritina x 0,7502; sideremia x 1,2574", "Fattori dedotti",
  "Il metodo BRINDA corregge i marcatori marziali in base ai reattanti di fase acuta. I marcatori di infiammazione non sono pubblicati in questo articolo, quindi i fattori sono stati ricavati dal rapporto fra valori aggiustati e grezzi. Quello sulla ferritina corrisponde a una riduzione del 25% esatto."),
 (A3, "Transferrina / log(ferritina)", "transferrina (g/L) / log10(ferritina (ug/L))", "3,119 / 1,860 = 1,677 (referto 1,63)", "Ricostruita con riserva",
  "La forma della formula e certa, il valore no: dipende dai pesi molecolari usati per convertire transferrina e ferritina dalle unita SI a quelle convenzionali, che la fonte non dichiara. Con i pesi plausibili il risultato oscilla fra 1,59 e 1,70 e l'1,63 pubblicato cade dentro la banda. Il modello usa 79.570 g/mol per la transferrina e 450.000 per la ferritina."),
 (A3, "Conversioni verso unita convenzionali", "ferro: uM x 5,585 = ug/dL | ferritina: pM x 0,45 = ug/L | transferrina: uM x 7,957 = mg/dL", "sideremia 76 ug/dL; ferritina 72,4 ug/L", "Standard",
  "Il pannello e espresso in unita SI, che quasi nessun laboratorio italiano usa. Il foglio 09 riporta la conversione completa per rendere i valori confrontabili con un referto ordinario."),
 (A4, "Neutrofili / Linfociti", "NLR = neutrofili / linfociti", "5,868 / 1,741 = 3,371 (referto 3,4)", "Esplicita",
  "Essendo un rapporto si puo calcolare anche sulle percentuali, perche il totale dei leucociti si semplifica: 69/20,5 da 3,366, praticamente lo stesso numero."),
 (A4, "Monociti / Linfociti", "MLR = monociti / linfociti", "0,6615 / 1,7407 = 0,380", "Esplicita", "Riproduce esattamente il valore pubblicato."),
 (A4, "Piastrine / Linfociti", "PLR = piastrine / linfociti", "272 / 1,7407 = 156,3 (referto 158)", "Esplicita",
  "Scarto dell'1,1%, dovuto agli arrotondamenti dell'emocromo ricostruito."),
 (A4, "Systemic Inflammation Index", "SII = piastrine x neutrofili / linfociti", "272 x 5,868 / 1,7407 = 917,0", "Esplicita",
  "Riproduce esattamente il valore pubblicato. E il prodotto delle tre popolazioni piu informative in un solo numero."),
 (A4, "Systemic Inflammation Response Index", "SIRI = neutrofili x monociti / linfociti", "5,868 x 0,6615 / 1,7407 = 2,230", "Esplicita",
  "Riproduce esattamente il valore pubblicato."),
 (A4, "Aggregate Index of Systemic Inflammation", "AISI = neutrofili x monociti x piastrine / linfociti", "SIRI x piastrine = 606,56", "Esplicita",
  "Riproduce esattamente il valore pubblicato. Nota utile: AISI diviso SIRI restituisce direttamente il numero delle piastrine, ed e la relazione da cui parte la ricostruzione dell'emocromo nel foglio 01."),
 (A4, "Monociti / HDL", "MHR = monociti (10^9/L) / HDL (mg/dL) x 1000", "0,6615 / 41 x 1000 = 16,13 (referto 16,1)", "Ricostruita",
  "L'espressione per mille con l'HDL in mg/dL riproduce il valore pubblicato. La forma piu diffusa in letteratura usa l'HDL in mmol/L e darebbe 0,62, un numero senza rapporto con il cut-off di 6."),
 (A4, "Uricemia / HDL", "UHR = uricemia / HDL x 100", "6,7 / 41 x 100 = 16,34 (referto 16,3)", "Ricostruita",
  "Riproduce il valore pubblicato. Entrambi i termini in mg/dL."),
 (A4, "MPV / Piastrine", "MPV / piastrine x 100", "11,906 / 272 x 100 = 4,38 (referto 4,3)", "Ricostruita",
  "Il fattore 100 e necessario per ottenere l'ordine di grandezza pubblicato. Scarto dell'1,8%."),
 (A4, "Emoglobina / RDW", "Hb (g/dL) / RDW (%)", "13,9 / 14,5 = 0,959 -> 0,96", "Ricostruita con riserva",
  "La forma e certa, i due valori no: l'articolo non pubblica ne emoglobina ne RDW, e il sistema resta sottodeterminato perche RDW/Piastrine ha due sole cifre decimali. Sono stati scelti i valori coerenti con l'anemia lieve descritta dalla serie."),
 (A4, "Ricostruzione dell'emocromo", "piastrine = AISI/SIRI; monociti = AISI/SII; linfociti = monociti/MLR; neutrofili = SIRI x linfociti / monociti", "P 272; M 0,662; L 1,741; N 5,868", "Inversione degli indici",
  "L'articolo non pubblica i conteggi assoluti, ma il sistema di sei indici in quattro incognite e sovradeterminato e si risolve. La verifica incrociata restituisce SII, SIRI, AISI e MLR esatti, NLR entro lo 0,8% e PLR entro l'1,1%."),
]
r = 5
for i, (fonte, nome, f, ver, orig, nota) in enumerate(FORM, 1):
    ws.cell(row=r, column=1, value=i).alignment = Alignment(horizontal="center", vertical="top")
    ws.cell(row=r, column=2, value=fonte.split(" - ")[0]).font = Font(name=FONT, size=9, color="595959")
    ws.cell(row=r, column=3, value=nome).font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=4, value=f).font = Font(name=FONT, size=9, color="C00000")
    ws.cell(row=r, column=5, value=ver).font = Font(name=FONT, size=9)
    bad = orig.startswith("Non riproducibile")
    ws.cell(row=r, column=6, value=orig).font = Font(
        name=FONT, size=9, bold=(orig == "Ricostruita" or bad),
        color="C00000" if bad else ("7030A0" if orig == "Ricostruita" else "595959"))
    cn = ws.cell(row=r, column=7, value=nota)
    cn.font = Font(name=FONT, size=9)
    ws.row_dimensions[r].height = rowh(nota, 82, 28, 78)
    for cc in range(1, 8):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="top",
                                   horizontal="center" if cc == 1 else "left")
    if bad:
        for cc in range(1, 8):
            ws.cell(row=r, column=cc).fill = ERR_FILL
    r += 1
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:G{r-1}"

# =====================================================================
# 04 ZONE FC
# =====================================================================
ws = wb.create_sheet("04 Zone FC")
setw(ws, {"A": 8, "B": 26, "C": 12, "D": 12, "E": 12, "F": 14, "G": 62})
title(ws, "04 - Zone di frequenza cardiaca",
      "Ricavate dalla FC massima del foglio 01. Le soglie si ottengono troncando all'intero la percentuale della FC massima.")
hdr(ws, 4, ["Zona", "Obiettivo", "% min", "% max", "Da (bpm)", "A (bpm)", "Nota"])
ZONES = [(0, "Riposo", 0.0, 0.5, "Zona in cui si trova il soggetto a riposo."),
         (1, "Attivita moderata", 0.5, 0.6, "Il limite inferiore e il confine che il soggetto sfiora da fermo."),
         (2, "Controllo del peso", 0.6, 0.7, ""),
         (3, "Aerobiosi", 0.7, 0.8, ""),
         (4, "Anaerobiosi", 0.8, 0.9, ""),
         (5, "Massimale", 0.9, 1.0, "Da leggere con la tolleranza del 90-110% sulla FC massima predetta.")]
r = 5
for z, obj, pmin, pmax, nota in ZONES:
    ws.cell(row=r, column=1, value=z).font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=2, value=obj)
    ws.cell(row=r, column=3, value=pmin).number_format = "0%"
    ws.cell(row=r, column=4, value=pmax).number_format = "0%"
    ws.cell(row=r, column=5, value=0 if z == 0 else f"=INT(C{r}*{I_HRMAX})")
    ws.cell(row=r, column=6, value=f"={I_HRMAX}" if z == 5 else f"=INT(C{r+1}*{I_HRMAX})-1")
    ws.cell(row=r, column=7, value=nota).font = Font(name=FONT, size=9)
    for cc in range(1, 8):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        cell.alignment = Alignment(horizontal="center" if cc in (1, 3, 4, 5, 6) else "left",
                                   vertical="center", wrap_text=(cc == 7))
    ws.row_dimensions[r].height = 26
    r += 1
r += 1
for k, (lab, val, col) in enumerate([
        ("FC a riposo del soggetto", f"={I_HR}", GREEN),
        ("Soglia di uscita dalla zona 0", "=E6", BLACK),
        ("Margine residuo prima di uscire dal riposo", f"=E{r+1}-E{r}", "C00000")]):
    ws.cell(row=r + k, column=2, value=lab).font = Font(name=FONT, size=10, bold=True)
    c = ws.cell(row=r + k, column=5, value=val)
    c.font = Font(name=FONT, size=11 if k == 2 else 10, bold=True, color=col)
    c.alignment = Alignment(horizontal="center")
    for cc in range(2, 8):
        ws.cell(row=r + k, column=cc).border = BORDER
ws.cell(row=r + 2, column=7, value="Battiti che separano il soggetto, da fermo, dalla zona dell'attivita moderata.").font = Font(name=FONT, size=9, italic=True)

# =====================================================================
# 05 EVIDENZE
# =====================================================================
ws = wb.create_sheet("05 Evidenze")
setw(ws, {"A": 5, "B": 11, "C": 12, "D": 34, "E": 24, "F": 26, "G": 16, "H": 22, "I": 56})
title(ws, "05 - Evidenze citate dalle fonti",
      "Ogni riga riporta la misura di rischio come pubblicata e la sua applicabilita al caso in esame.")
hdr(ws, 4, ["#", "Fonte", "PMID", "Studio / disegno", "Numerosita", "Esposizione", "Misura", "Esito", "Applicabilita al caso"])
EV = [
 (A1, "26598376", "Meta-analisi di 46 studi prospettici", "1.246.203 soggetti / 78.349 decessi", "FC a riposo, per ogni +10 bpm", "RR 1,09", "Mortalita totale",
  "Il soggetto e a 82 bpm: 17 bpm oltre la fascia di minimo, circa +16% di rischio relativo per estrapolazione."),
 (A1, "26598376", "Meta-analisi - analisi per categorie", "1.246.203 soggetti", "FC a riposo > 80 bpm", "RR 1,45", "Mortalita totale",
  "SOGLIA SUPERATA (82 bpm). E il salto di scala descritto dalla fonte."),
 (A1, "26598376", "Meta-analisi - analisi per categorie", "1.246.203 soggetti", "FC a riposo > 80 bpm", "RR 1,33", "Mortalita cardiovascolare", "Soglia superata."),
 (A1, "11337213", "Coorte giapponese, 18 anni", "573 uomini, 40-64 anni", "FC > 90 vs 60-69 bpm", "RR 2,68", "Mortalita",
  "Il minimo di mortalita cade fra 60 e 69 bpm e l'aumento comincia ben prima dei 90: il soggetto e nel tratto ascendente."),
 (A1, "28067310", "Coorte di Kailuan, 3 misurazioni in 4 anni", "47.311 adulti", "FC >= 80 bpm in tutte e tre le misurazioni", "HR 1,86", "Mortalita per ogni causa",
  "Lo studio piu pertinente perche misura la persistenza. Applicabile solo se il valore viene confermato nel tempo: servono misurazioni ripetute."),
 (A1, "17446799", "Coorte prospettica, 5 anni", "1.910 uomini", "Riserva impegnata sotto l'80%", "HR 2,8", "Mortalita cardiovascolare",
  "NON VALUTABILE: richiede un test da sforzo che il referto non contiene. E il principale accertamento mancante."),
 (A1, "10353296", "Revisione (Palatini) sull'ipertono simpatico", "-", "Tachicardia cronica", "-", "Meccanismo",
  "Doppio meccanismo: marcatore di iperattivita adrenergica e stimolo meccanico diretto, con stress di parete che favorisce la placca."),
 (A1, "33793325", "Revisione su Circulation Research", "-", "Pressione di pulsazione elevata", "-", "Invecchiamento vascolare",
  "La pulsatilita si scarica su cervello e rene e riduce la diastolica utile alla perfusione coronarica. Il soggetto e dentro il range."),
 (A1, "38453019", "Letteratura sul rate pressure product", "-", "Prodotto cardiovascolare", "-", "Consumo miocardico di O2",
  "Valida il prodotto cardiovascolare come surrogato non invasivo del consumo miocardico di ossigeno."),
 (A1, "35328926", "Effetti acuti del fumo", "giovani normotesi", "Singola sigaretta", "-", "Aumento acuto di MAP e FC",
  "Effetto acuto: su un referto a riposo conta poco. Rileva cio che resta fra una sigaretta e l'altra."),
 (A1, "8689656", "Caso-controllo su variabilita RR", "20 fumatori > 20 sig/die vs controlli", "Fumo abituale", "LF 70,6 vs 46,0; HF 22,1 vs 42,0", "Assetto autonomico spostato",
  "APPLICABILE: il soggetto fuma 20 sigarette al giorno. Piu acceleratore simpatico, meno freno vagale. Parte dei battiti in eccesso viene da qui."),
 (A1, "16029383", "Studio su forti fumatori", "-", "Durata dell'abitudine", "-", "Calo della modulazione vagale",
  "APPLICABILE E PIU PERTINENTE: misura gli anni, non solo le sigarette. Il soggetto ne ha 30."),
 (A1, "15301332", "Letteratura su apnee ostruttive del sonno", "-", "OSAS", "-", "FC a riposo alta, R-R ridotta, variabilita pressoria",
  "SOSPETTO NON INDAGATO: russamento e pause riferite, nessun accertamento. Spiegherebbe parte della frequenza elevata."),
 (A1, "12609010", "Traffico nervoso simpatico", "-", "OSAS", "-", "Simpatico elevato anche in veglia",
  "L'effetto non si esaurisce con il risveglio: persiste con respiro normale e senza ipossia."),
 (A1, "15793048", "Randomizzato controllato", "136 uomini con coronaropatia", "3 mesi di esercizio supervisionato", "-19% RPP", "Prodotto cardiovascolare a riposo",
  "INTERVENTO EFFICACE. Nel gruppo non supervisionato il valore e rimasto invariato: la supervisione e parte dell'effetto. Simulato nel foglio 07."),
 (A1, "8651120", "Coorte di cessazione tabagica", "54 fumatori da >= 1 pacchetto/die", "Cessazione", "-", "Calo FC, miglioramento della variabilita",
  "EFFICACE ma non risolutivo nel breve: a 4 settimane la FC restava sopra e la variabilita sotto i valori attesi per l'eta."),
 (A1, "11153730", "Meta-analisi su FC massima", "351 studi / 18.712 soggetti", "Eta", "FCmax = 208 - 0,7 x eta", "Stima della FC massima",
  "Termine di confronto con il modello logistico del referto."),
 (A2, "22815707", "Coorte NHANES (Krakauer)", "14.105 adulti", "Forma corporea depurata da peso e BMI", "aumento quasi esponenziale", "Mortalita",
  "APPLICABILE E IL PIU GRAVE: z-score 1,52 contro un limite di 0,228. E l'indice che un normopeso non si aspetta mai di trovare alterato."),
 (A2, "22106927", "Meta-analisi (Ashwell)", "31 studi", "Rapporto vita/altezza", "-", "Rischio cardiometabolico",
  "APPLICABILE: WHtR 0,56 contro 0,52. Il rapporto vita/altezza discrimina meglio del BMI. Un metro da sarta batte la bilancia."),
 (A2, "16150143", "Studio di validazione (Kahn)", "-", "Lipid Accumulation Product", "-", "Rischio cardiovascolare",
  "APPLICABILE: LAP 53,5 contro 26,7, il doppio esatto del cut-off. Riconosce il rischio meglio del BMI."),
 (A2, "20067971", "Validazione (Amato)", "-", "Visceral Adiposity Index", "-", "Funzione del tessuto adiposo",
  "APPLICABILE: VAI 2,29 contro 1,92. Marcatore della funzione del grasso viscerale, non della sua quantita."),
 (A2, "25199852", "Costruzione dell'indice (Wakabayashi e Daimon)", "-", "Cardiometabolic Index", "-", "Diabete e disglicemia",
  "APPLICABILE: CMI 0,88 contro 0,39, piu del doppio. Buon discriminatore di disglicemia."),
 (A2, "26551006", "Coorte (Sahakyan)", "-", "Obesita centrale in soggetti normopeso", "-", "Mortalita cardiovascolare",
  "DESCRIVE ESATTAMENTE IL CASO: nel normopeso l'obesita centrale si accompagna a mortalita superiore a quella dell'obeso con distribuzione periferica. Il magro con il girovita largo e il caso peggiore, non il migliore."),
 (A2, "30312372", "Consenso europeo EWGSOP2", "-", "Circonferenza del polpaccio", "-", "Massa muscolare appendicolare",
  "APPLICABILE: valida il polpaccio come surrogato della massa muscolare. La correzione per BMI usata dal referto non e pero quella del consenso: vedi foglio 08."),
 (A2, "35227529", "Consenso ESPEN-EASO (Donini)", "-", "Obesita sarcopenica", "-", "Insulino-resistenza, diabete 2, ridotta mobilita",
  "APPLICABILE: il soggetto non e obeso secondo il BMI ma ha gia il fenotipo. Formalizza l'associazione fra infiltrazione adiposa del muscolo e insulino-resistenza."),
 (A3, "35623855", "Studio su donne in eta fertile (Finkelstein)", "979 donne", "Correzione della ferritina per infiammazione (BRINDA)", "46,3% -> 61,5%", "Prevalenza di carenza marziale",
  "APPLICABILE: una persona su sette veniva dichiarata sana per un errore di lettura. Il soggetto e un fumatore da trent'anni, quindi con flogosi cronica di basso grado: e esattamente la condizione in cui la ferritina grezza inganna."),
 (A3, "27346617", "Revisione (Elsayed, Sharif, Stack)", "-", "Saturazione della transferrina", "-", "Stato marziale reale",
  "APPLICABILE: la saturazione e il marcatore piu informativo perche combina ferro circolante e capacita di trasporto. Nel soggetto vale 17,3% contro un minimo di 20."),
 (A3, "36380788", "Revisione (Del Pinto, Ferri)", "-", "Carenza marziale con infiammazione cronica", "-", "Ruolo discriminante della saturazione",
  "APPLICABILE: riconosce alla saturazione un ruolo discriminante proprio nelle situazioni in cui la ferritina non e leggibile."),
 (A3, "22306005", "Revisione (Ganz e Nemeth)", "-", "Epcidina e ferroportina", "-", "Meccanismo di regolazione del ferro",
  "IPOTESI, NON MISURA: spiega il meccanismo con cui il ferro resta bloccato nei depositi, ma l'epcidina non e stata dosata e non e dosabile nella routine clinica."),
 (A3, "30401705", "Revisione (Weiss, Ganz, Goodnough)", "-", "Anemia da infiammazione", "-", "Eritropoiesi ferro-ristretta",
  "APPLICABILE E DECISIVA SULLA CONDOTTA: nelle anemie da infiammazione la sola supplementazione marziale e spesso inefficace e la strategia utile passa dal trattare la causa."),
 (A3, "27236129", "Revisione (Bahrainwala, Berns)", "-", "Ferritina e saturazione prese da sole", "-", "Capacita predittiva modesta",
  "APPLICABILE: giustifica il commento 'profilo marziale non discriminante'. E il fondamento del rifiuto dell'algoritmo a dare una risposta che i dati non contengono."),
 (A3, "31850722", "Studio (Gelaw, Woldu, Melku)", "-", "Marcatori dello stato marziale", "-", "Valutazione della carenza",
  "Citata dalla fonte fra i riferimenti senza essere discussa nel corpo dell'articolo."),
 (A4, "36769776", "Coorte NHANES, 20 anni di follow-up (Xia)", "42.875 adulti", "SII e SIRI elevati", "associazione significativa", "Mortalita totale e cardiovascolare",
  "APPLICABILE A ENTRAMBI: il soggetto ha SII appena sopra il limite e SIRI a tre volte e mezzo la soglia."),
 (A4, "37265570", "Coorte di ipertesi (Xiu)", "23.765 partecipanti", "AISI, quartile piu alto vs piu basso", "HR 1,91", "Mortalita cardiovascolare",
  "APPLICABILE: l'AISI del soggetto e quattro volte il cut-off, quindi ampiamente nel quartile superiore. E la misura di rischio piu forte fra quelle citate dalle quattro fonti."),
 (A4, "38755603", "Studio su indicatori derivati dall'emocromo", "-", "Indici infiammatori da emocromo", "-", "Sarcopenia e mortalita",
  "CHIUDE UN CERCHIO: collega gli indici di questa sezione alla sarcopenia documentata nell'articolo 2, dove il polpaccio corretto risultava 33,5 cm contro una soglia di 34,0."),
 (A4, "34109496", "Studio italiano su nefropatici (Gembillo)", "214 pazienti", "Rapporto monociti / HDL", "correlazione", "PCR, conta leucocitaria, ipertensione resistente",
  "APPLICABILE: il rapporto del soggetto e quasi il triplo del limite. Notevole che risulti correlato alla PCR proprio in un caso in cui la PCR e normale."),
 (A4, "30353545", "Revisione sul valore prognostico del SII", "-", "Systemic Inflammation Index", "-", "Prognosi nei tumori gastrointestinali",
  "Contesto oncologico, non direttamente trasferibile al caso: sostiene la validita dell'indice, non la sua interpretazione qui."),
 (A4, "31148950", "Revisione (Korniluk)", "-", "Volume piastrinico medio", "-", "Attivazione piastrinica nell'infiammazione",
  "APPLICABILE: sostiene i due rapporti basati sull'MPV, entrambi oltre soglia nel soggetto."),
 (A4, "12160596", "Studio metodologico (Van Tiel)", "-", "VES", "-", "Limiti interpretativi",
  "Sostiene il ridimensionamento della VES come strumento per l'infiammazione cronica di basso grado."),
]
r = 5
for i, (fonte, pmid, studio, n, esp, mis, esito, appl) in enumerate(EV, 1):
    ws.cell(row=r, column=1, value=i)
    ws.cell(row=r, column=2, value=fonte.split(" - ")[0]).font = Font(name=FONT, size=9, color="595959")
    cp = ws.cell(row=r, column=3, value=pmid)
    cp.font = Font(name=FONT, size=10, color="0563C1", underline="single")
    cp.hyperlink = f"https://pubmed.ncbi.nlm.nih.gov/{pmid}/"
    ws.cell(row=r, column=4, value=studio)
    ws.cell(row=r, column=5, value=n)
    ws.cell(row=r, column=6, value=esp)
    ws.cell(row=r, column=7, value=mis).font = Font(name=FONT, size=9, bold=True, color="C00000")
    ws.cell(row=r, column=8, value=esito)
    ws.cell(row=r, column=9, value=appl)
    for cc in range(1, 10):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=9)
        cell.alignment = Alignment(wrap_text=True, vertical="top",
                                   horizontal="center" if cc in (1, 2, 3, 7) else "left")
    ws.row_dimensions[r].height = rowh(appl, 58, 28, 70)
    r += 1
ws.freeze_panes = "D5"
ws.auto_filter.ref = f"A4:I{r-1}"

# =====================================================================
# 06 RISCHIO
# =====================================================================
ws = wb.create_sheet("06 Rischio")
setw(ws, {"A": 4, "B": 54, "C": 16, "D": 12, "E": 12, "F": 10, "G": 66})
title(ws, "06 - Valutazione della pericolosita",
      "Tre letture indipendenti: il punteggio strumentale dagli indici delle due fonti, i fattori anamnestici che i referti non contengono e il rischio relativo epidemiologico.")

def section(row, text):
    ws.cell(row=row, column=2, value=text).font = Font(name=FONT, size=11, bold=True, color="FFFFFF")
    for cc in range(2, 8):
        ws.cell(row=row, column=cc).fill = HDR_FILL
        ws.cell(row=row, column=cc).border = BORDER
    ws.row_dimensions[row].height = 20

def line(row, label, formula, fmt="0.00", note="", bold=False, color=BLACK, size=10):
    ws.cell(row=row, column=2, value=label).font = Font(name=FONT, size=size, bold=bold)
    c = ws.cell(row=row, column=3, value=formula)
    c.font = Font(name=FONT, size=size, bold=True, color=color)
    c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    n = ws.cell(row=row, column=7, value=note)
    n.font = Font(name=FONT, size=9, italic=True, color="595959")
    n.alignment = Alignment(wrap_text=True, vertical="center")
    for cc in range(2, 8):
        ws.cell(row=row, column=cc).border = BORDER
    if note:
        ws.row_dimensions[row].height = rowh(note, 82, 15, 40)

M = f"'02 Indici'!$M${R0}:$M${LAST}"
Q = f"'02 Indici'!$Q${R0}:$Q${LAST}"
P = f"'02 Indici'!$P${R0}:$P${LAST}"
N = f"'02 Indici'!$N${R0}:$N${LAST}"
O = f"'02 Indici'!$O${R0}:$O${LAST}"

section(4, "A - PUNTEGGIO STRUMENTALE (59 indici, quattro fonti)")
line(5, "Indici valutati in totale", f"=COUNT({O})", "0", "Righe presenti nel foglio 02.")
line(6, "Indici normali", f'=COUNTIF({M},"Normale")', "0")
line(7, "Indici borderline (in zona di guardia)", f'=COUNTIF({M},"Borderline")', "0",
     "Dentro l'intervallo ma nell'ultimo 20% verso il lato critico. I referti li stampano come normali.")
line(8, "Indici fuori intervallo", f'=COUNTIF({M},"Fuori (alto)")+COUNTIF({M},"Fuori (basso)")', "0", "", bold=True, color="C00000")
line(9, "  di cui dall'articolo 1 (cardiologici)", f'=COUNTIFS({Q},"Art.1*",{M},"Fuori (alto)")+COUNTIFS({Q},"Art.1*",{M},"Fuori (basso)")', "0",
     "La fonte dichiara 1 indice fuori range su 16.")
line(10, "  di cui dall'articolo 2 (antropometrici)", f'=COUNTIFS({Q},"Art.2*",{M},"Fuori (alto)")+COUNTIFS({Q},"Art.2*",{M},"Fuori (basso)")', "0",
     "La fonte dichiara 6 indici oltre soglia su 12; applicando i cut-off delle sue stesse tabelle ne risultano 9. Vedi foglio 08.")
line(11, "  di cui dall'articolo 3 (profilo marziale)", f'=COUNTIFS({Q},"Art.3*",{M},"Fuori (alto)")+COUNTIFS({Q},"Art.3*",{M},"Fuori (basso)")', "0",
     "La fonte dichiara 5 valori fuori dalle finestre ottimali su 11; applicando gli intervalli della sua stessa tabella ne risultano 6. Vedi foglio 08.")
line(12, "  di cui dall'articolo 4 (infiammatori)", f'=COUNTIFS({Q},"Art.4*",{M},"Fuori (alto)")+COUNTIFS({Q},"Art.4*",{M},"Fuori (basso)")', "0",
     "La fonte dichiara 10 indici fuori riferimento su 16. Il modello ne trova 9 piu il rapporto monociti/linfociti esattamente sul limite superiore, che la fonte conta e il modello classifica come borderline.")
line(13, "Punteggio ponderato ottenuto", f"=SUM({P})", "0", "Somma di peso clinico x punteggio 0-3.")
line(14, "Punteggio massimo teorico", f"=SUM({N})*3", "0", "Se ogni indice pesato fosse gravemente fuori range.")
line(15, "COMPONENTE STRUMENTALE", "=IF(C14=0,0,C13/C14)", "0.0%", "Quota del massimo teorico raggiunta.", bold=True, color="C00000")

section(17, "B - FATTORI ANAMNESTICI E PATTERN CHE I SINGOLI INDICI NON ESPRIMONO")
hdr(ws, 18, ["", "Fattore", "Presente", "Peso", "Punti", "", "Evidenza / nota"], start=1, h=20)
FATT = [
 ("Fumo attivo (>= 20 sigarette al giorno)", f"=IF({I_CIG}>=20,1,0)", 3,
  "PMID 8689656 - assetto autonomico spostato verso il simpatico anche a distanza dall'ultima sigaretta."),
 ("Durata dell'abitudine tabagica >= 20 anni", f"=IF({I_YRS}>=20,1,0)", 2,
  "PMID 16029383 - la modulazione vagale cala al crescere degli anni, non solo delle sigarette."),
 ("Fenotipo TOFI (normopeso con obesita centrale)", f"=IF(AND({E_BMI}<25,{E_WHtR}>0.5),1,0)", 3,
  "PMID 26551006 - calcolato dagli indici: BMI sotto 25 con vita oltre meta dell'altezza. La mortalita cardiovascolare supera quella dell'obeso con distribuzione periferica."),
 ("Miosteatosi documentata", f"={I_MIO}", 2,
  "PMID 35227529 - infiltrazione adiposa del muscolo, non misurabile con l'antropometria di superficie. Un quadricipite infiltrato pesa uguale e lavora peggio."),
 ("Sospetta apnea ostruttiva del sonno mai indagata", f"={I_OSAS}", 3,
  "PMID 15301332 e 12609010 - il tono simpatico resta elevato anche in veglia. E il sospetto piu pesante proprio perche non e mai stato verificato."),
 ("Sedentarieta", f"={I_SED}", 1,
  "PMID 15793048 - l'esercizio supervisionato abbassa del 19% il prodotto cardiovascolare; la sua assenza toglie il principale fattore correttivo."),
 ("Incremento ponderale > 5 kg dall'eta di 30 anni", f"=IF({I_W}-{I_W30}>5,1,0)", 1,
  "Sei chili in vent'anni, trecento grammi l'anno: l'aumento piu banale del mondo. Sotto, circa dieci chili di grasso in piu e quattro di muscolo in meno."),
 ("Pattern di eritropoiesi ferro-ristretta", f"=IF(AND({E_SAT}<20,{E_DEP}>301.5),1,0)", 3,
  "PMID 30401705 - calcolato dagli indici: saturazione sotto il 20% con depositi non esauriti. Il ferro c'e ma non arriva ai reticolociti. Nessuna singola riga del profilo marziale lo dice: emerge solo dalla combinazione."),
 ("Infiammazione cronica di basso grado con indici classici negativi",
  f'=IF(AND({E_PCR}<=0.5,COUNTIFS({Q},"Art.4*",{M},"Fuori (alto)")+COUNTIFS({Q},"Art.4*",{M},"Fuori (basso)")>=5),1,0)', 3,
  "PMID 36769776 e 37265570 - calcolato: PCR normale con almeno cinque indici derivati oltre soglia. E la condizione che rende non verificabile l'aggiustamento per infiammazione dell'articolo 3 e che al tempo stesso lo giustifica sul piano fisiopatologico."),
]
r = 19
f0 = r
for lab, presf, peso, ev in FATT:
    ws.cell(row=r, column=2, value=lab).font = Font(name=FONT, size=10)
    c = ws.cell(row=r, column=3, value=presf)
    c.font = Font(name=FONT, size=10, bold=True)
    c.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=4, value=peso).font = Font(name=FONT, size=10, color=BLUE)
    ws.cell(row=r, column=4).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=5, value=f"=C{r}*D{r}").alignment = Alignment(horizontal="center")
    n = ws.cell(row=r, column=7, value=ev)
    n.font = Font(name=FONT, size=9, italic=True, color="595959")
    n.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = rowh(ev, 82, 26, 44)
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).border = BORDER
        if ws.cell(row=r, column=cc).font.name != FONT:
            ws.cell(row=r, column=cc).font = Font(name=FONT, size=10)
    r += 1
f1 = r - 1
line(r, "Punti ottenuti", f"=SUMPRODUCT(C{f0}:C{f1},D{f0}:D{f1})", "0", "", bold=True)
line(r + 1, "Punti massimi", f"=SUM(D{f0}:D{f1})", "0", "", bold=True)
line(r + 2, "COMPONENTE ANAMNESTICA", f"=IF(C{r+1}=0,0,C{r}/C{r+1})", "0.0%",
     "La sarcopenia non compare qui perche e gia misurata fra gli indici del foglio 02, per evitare di contarla due volte.",
     bold=True, color="C00000")
FATT_PCT = f"C{r+2}"
FATT_GOT, FATT_MAX = f"C{r}", f"C{r+1}"

er = r + 4
section(er, "C - RISCHIO EPIDEMIOLOGICO CALCOLATO")
line(er + 1, "Scostamento dalla FC di minimo rischio (60-69 bpm)", f"={I_HR}-{I_HROPT}", "0", "bpm oltre il punto medio della fascia con mortalita minima.")
line(er + 2, "Scostamento dal limite superiore del referto", f"={I_HR}-{I_HRREF}", "0", "bpm oltre il valore che il referto considera normale.")
line(er + 3, "RR mortalita totale (+1,09 ogni 10 bpm)", f"=1.09^(({I_HR}-{I_HROPT})/10)", "0.000", "PMID 26598376. Stima di gradiente, non rischio assoluto individuale.")
line(er + 4, "RR mortalita cardiovascolare (+1,08 ogni 10 bpm)", f"=1.08^(({I_HR}-{I_HROPT})/10)", "0.000", "PMID 26598376.")
line(er + 5, "Soglia FC > 80 bpm superata", f'=IF({I_HR}>{I_HRTHR},"SI","NO")', "General",
     "Oltre gli 80 bpm il quadro cambia scala: RR 1,45 totale e 1,33 cardiovascolare.", color="C00000")
line(er + 6, "RR categoriale - mortalita totale", f'=IF({I_HR}>{I_HRTHR},1.45,"n.a.")', "0.00", "PMID 26598376.")
line(er + 7, "RR categoriale - mortalita cardiovascolare", f'=IF({I_HR}>{I_HRTHR},1.33,"n.a.")', "0.00", "PMID 26598376.")
line(er + 8, "HR se il valore >= 80 bpm si conferma nel tempo", f'=IF({I_HR}>={I_HRTHR},1.86,"n.a.")', "0.00",
     "PMID 28067310. Richiede misurazioni ripetute, oggi non disponibili.")
line(er + 9, "Battiti in eccesso all'anno", f"=({I_HR}-{I_HRREF})*60*24*365", "#,##0",
     "Il carico cumulativo che nessun sintomo segnala e nessuna visita registra.", bold=True)
line(er + 10, "Battiti in eccesso in 10 anni", f"=({I_HR}-{I_HRREF})*60*24*365*10", "#,##0")
line(er + 11, "Criterio vita per sindrome metabolica", f'=IF({I_WC}>={I_WCMS},"SODDISFATTO","NON soddisfatto")', "General",
     "Il criterio binario resta vuoto per 4 cm, mentre vita/altezza, conicita e ABSI sono gia tutti oltre i rispettivi cut-off. E il limite delle soglie a interruttore su variabili continue.")
line(er + 12, "Centimetri di vita che mancano al criterio", f"={I_WCMS}-{I_WC}", "0.0", "")

cr = er + 14
section(cr, "D - INDICE COMPOSITO DI PERICOLOSITA")
c = ws.cell(row=cr + 1, column=3, value=0.6)
line(cr + 1, "Peso della componente strumentale", 0.6, "0%", "Modificabile: le due componenti devono sommare a 1.")
ws.cell(row=cr + 1, column=3).font = Font(name=FONT, size=10, bold=True, color=BLUE)
ws.cell(row=cr + 1, column=3).fill = PatternFill("solid", fgColor=YEL)
line(cr + 2, "Peso della componente anamnestica", f"=1-C{cr+1}", "0%",
     "I referti coprono la maggior parte del giudizio, ma non tutto: il resto viene da cio che non misurano.")
line(cr + 3, "PUNTEGGIO DI PERICOLOSITA (0-100)", f"=(C15*C{cr+1}+{FATT_PCT}*C{cr+2})*100", "0.0",
     "Media pesata delle due componenti su scala centesimale.", bold=True, color="C00000", size=11)
ws.cell(row=cr + 3, column=3).font = Font(name=FONT, size=18, bold=True, color="C00000")
ws.row_dimensions[cr + 3].height = 30
line(cr + 4, "FASCIA", f'=IF(C{cr+3}<20,"BASSO",IF(C{cr+3}<40,"LIEVE",IF(C{cr+3}<60,"MODERATO",IF(C{cr+3}<80,"ELEVATO","ALTO"))))',
     "General", "Fasce: <20 basso | 20-40 lieve | 40-60 moderato | 60-80 elevato | >=80 alto.", bold=True, size=12)
ws.cell(row=cr + 4, column=3).font = Font(name=FONT, size=15, bold=True, color="C00000")
ws.row_dimensions[cr + 4].height = 26

ws.cell(row=cr + 6, column=2, value="Lettura sintetica").font = Font(name=FONT, size=10, bold=True)
c = ws.cell(row=cr + 6, column=3, value=(
    f'="Su "&C5&" indici ricalcolati, "&C8&" risultano fuori intervallo nella direzione critica e "&C7&'
    f'" sono in zona di guardia pur essendo stampati come normali. "'
    f'&IF({I_HR}>{I_HRTHR},"La frequenza a riposo supera la soglia degli 80 bpm, oltre la quale il rischio cambia scala. ","")'
    f'&IF(AND({E_BMI}<25,{E_WHtR}>0.5),"Il BMI resta sotto 25 mentre la vita supera la meta dell altezza: e il fenotipo che il peso corporeo nasconde. ","")'
    f'&IF(AND({E_SAT}<20,{E_DEP}>301.5),"Il profilo marziale mostra saturazione bassa con depositi non esauriti: il ferro c e ma non arriva dove serve, quindi la sola supplementazione marziale sarebbe verosimilmente inefficace e la leva utile e la causa infiammatoria. ","")'
    f'&IF({E_PCR}<=0.5,"I due indici infiammatori classici sono negativi, ma gli indici derivati dall emocromo documentano infiammazione cronica di basso grado: SIRI a "&TEXT({E_SIRI},"0,00")&" contro 0,68 e AISI a "&TEXT({E_AISI},"0")&" contro 147. ","")'
    f'&"Fattori anamnestici e pattern presenti: "&{FATT_GOT}&" punti su "&{FATT_MAX}&". "'
    f'&"Accertamenti mancanti che cambierebbero il giudizio: test da sforzo, studio del sonno, misurazioni ripetute della frequenza, marcatori di infiammazione e contenuto emoglobinico reticolocitario."'))
c.font = Font(name=FONT, size=10)
c.alignment = Alignment(wrap_text=True, vertical="top")
c.fill = WARN_FILL
ws.merge_cells(start_row=cr + 6, start_column=3, end_row=cr + 9, end_column=7)
for cc in range(3, 8):
    for rr in range(cr + 6, cr + 10):
        ws.cell(row=rr, column=cc).fill = WARN_FILL
        ws.cell(row=rr, column=cc).border = BORDER

mr = cr + 11
section(mr, "E - COSA I REFERTI NON POSSONO DIRE")
MANCA = [
 ("Test da sforzo", "Senza di esso non e valutabile la riserva effettivamente impegnata, il parametro con HR 2,8 per mortalita cardiovascolare (PMID 17446799)."),
 ("Studio del sonno", "Russamento e pause respiratorie riferiti ma mai indagati. Le apnee spiegherebbero una quota della frequenza elevata (PMID 15301332)."),
 ("Misurazioni ripetute della FC a riposo", "L'HR 1,86 di Kailuan si applica alla persistenza del valore su piu misurazioni, non a un singolo prelievo (PMID 28067310)."),
 ("Variabilita della frequenza cardiaca (HRV)", "LF, HF e guadagno barocettivo sono i parametri su cui il fumo agisce, e nessuno di essi compare nei referti."),
 ("Glicemia, insulinemia, enzimi epatici", "VAI e CMI segnalano resistenza insulinica e disfunzione del tessuto adiposo, ma la conferma diretta richiede i marcatori di sensibilita insulinica e di steatosi."),
 ("DEXA o bioimpedenziometria", "Gli indici antropometrici stratificano il rischio a costo quasi nullo, ma non quantificano la composizione corporea: sarcopenia e miosteatosi restano stime."),
 ("Marcatori di infiammazione (PCR, AGP)", "Il referto applica al profilo marziale una correzione BRINDA che dipende dai reattanti di fase acuta, ma quei valori non compaiono. Senza di essi i fattori di aggiustamento restano dedotti, non verificati."),
 ("Contenuto emoglobinico reticolocitario (Ret-He)", "E l'esame con cui la fonte stessa propone di rivalutare il quadro fra tre mesi. Direbbe se il ferro arriva davvero ai precursori, che e la domanda che il profilo marziale lascia aperta."),
 ("Epcidina", "Spiegherebbe il blocco ipotizzato, ma non e dosabile nella routine: i metodi non sono armonizzati e mancano cut-off condivisi. Resta uno strumento di ricerca, quindi il meccanismo si deduce, non si legge (PMID 22306005)."),
]
r = mr + 1
for lab, why in MANCA:
    ws.cell(row=r, column=2, value=lab).font = Font(name=FONT, size=10, bold=True)
    c = ws.cell(row=r, column=3, value=why)
    c.font = Font(name=FONT, size=9)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=7)
    ws.row_dimensions[r].height = 30
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).border = BORDER
    r += 1

dr = r + 1
ws.cell(row=dr, column=2, value="AVVERTENZA").font = Font(name=FONT, size=10, bold=True, color="C00000")
c = ws.cell(row=dr, column=3, value=(
    "Il punteggio composito e una sintesi documentale costruita su pesi scelti dall'analista, non uno score clinico "
    "validato. Serve a rendere esplicito e discutibile un giudizio, non a sostituirlo. I pesi sono modificabili nella "
    "colonna 'Peso' del foglio 02 e nella colonna D della sezione B. Prima di usare i numeri, leggere il foglio "
    "'08 Controlli': sette valori pubblicati non si riproducono dai dati grezzi."))
c.font = Font(name=FONT, size=9, italic=True)
c.alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=dr, start_column=3, end_row=dr + 2, end_column=7)
for cc in range(2, 8):
    for rr in range(dr, dr + 3):
        ws.cell(row=rr, column=cc).fill = WARN_FILL
        ws.cell(row=rr, column=cc).border = BORDER

# =====================================================================
# 07 SCENARI
# =====================================================================
ws = wb.create_sheet("07 Scenari")
setw(ws, {"A": 4, "B": 38, "C": 10, "D": 10, "E": 10, "F": 11, "G": 10, "H": 10,
          "I": 13, "J": 14, "K": 10, "L": 50})
title(ws, "07 - Scenari di intervento",
      "Le celle gialle sono le leve dello scenario; il resto e calcolato. I valori sono proiezioni sull'ordine di grandezza degli effetti riportati dagli studi, non risultati misurati.")

ws.cell(row=4, column=2, value="TAVOLA 1 - Leve cardiovascolari").font = Font(name=FONT, size=11, bold=True, color="1F3864")
hdr(ws, 5, ["", "Scenario", "FC\n(bpm)", "SV\n(mL)", "SBP\n(mmHg)", "MAP\n(mmHg)", "CO\n(L/min)",
            "MSI", "RPP\n(bpm*mmHg)", "Battiti extra\n/anno", "RR\nmort.tot.", "Commento"], start=1)
SC = [
 ("Situazione attuale", 82, 66, 128, "Quindici indici cardiologici su sedici stampati come normali."),
 ("Esercizio supervisionato, 3 mesi", 72, 66, 124, "PMID 15793048: -19% sul prodotto cardiovascolare. Nel gruppo non supervisionato il valore era rimasto invariato."),
 ("Cessazione del fumo, oltre 4 settimane", 77, 66, 126, "PMID 8651120: la FC cala e la variabilita migliora, ma a 4 settimane resta sopra i valori attesi per l'eta."),
 ("Esercizio + cessazione", 70, 70, 122, "Le due leve insieme, con recupero parziale di gittata dovuto al condizionamento."),
 ("Pompa piu capiente a pari portata", 68, 80, 128, "Stessa portata ottenuta con gittata nel medio dell'intervallo: il confronto che mostra cosa nasconde una portata normale."),
 ("Obiettivo fascia di minimo rischio", 65, 80, 120, "Frequenza nella fascia 60-69 bpm, minimo di mortalita secondo PMID 11337213."),
]
r = 6
s0 = r
for i, (nome, hr, sv, sbp, comm) in enumerate(SC, 1):
    ws.cell(row=r, column=1, value=i).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=2, value=nome).font = Font(name=FONT, size=10, bold=(i == 1))
    for col, val in ((3, hr), (4, sv), (5, sbp)):
        c = ws.cell(row=r, column=col, value=val)
        c.font = Font(name=FONT, size=10, bold=True, color=BLUE)
        c.fill = PatternFill("solid", fgColor=YEL)
    ws.cell(row=r, column=6, value=f"={I_DBP}+0.4*(E{r}-{I_DBP})").number_format = "0"
    ws.cell(row=r, column=7, value=f"=D{r}*C{r}/1000").number_format = "0.0"
    ws.cell(row=r, column=8, value=f"=C{r}/F{r}").number_format = "0.00"
    ws.cell(row=r, column=9, value=f"=C{r}*E{r}").number_format = "#,##0"
    ws.cell(row=r, column=10, value=f"=MAX(0,C{r}-{I_HRREF})*60*24*365").number_format = "#,##0"
    ws.cell(row=r, column=11, value=f"=1.09^((C{r}-{I_HROPT})/10)").number_format = "0.000"
    ws.cell(row=r, column=12, value=comm).font = Font(name=FONT, size=9)
    for cc in range(1, 13):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        cell.alignment = Alignment(horizontal="center" if 3 <= cc <= 11 or cc == 1 else "left",
                                   vertical="center", wrap_text=(cc in (2, 12)))
    ws.row_dimensions[r].height = rowh(comm, 56, 28, 52)
    r += 1
s1 = r - 1

r += 1
ws.cell(row=r, column=2, value="TAVOLA 2 - Leve antropometriche").font = Font(name=FONT, size=11, bold=True, color="1F3864")
hdr(ws, r + 1, ["", "Scenario", "Peso\n(kg)", "Vita\n(cm)", "Fianchi\n(cm)", "BMI", "WHtR", "Conicity",
                "LAP", "CMI", "VAI", "Commento"], start=1)
SC2 = [
 ("Situazione attuale", 74.0, 97, 100, "BMI normale, sei indici di forma e ibridi oltre soglia."),
 ("-5 cm di vita a pari peso", 74.0, 92, 99, "Ricomposizione senza calo ponderale: e cio che l'esercizio contro resistenza produce per primo."),
 ("-3 kg con -6 cm di vita", 71.0, 91, 98, "Rientro nella forbice del peso desiderabile (68-71 kg secondo Lorenz e Creff)."),
 ("Vita sotto meta altezza (WHtR 0,50)", 71.0, 87, 97, "Soglia della regola aurea: la vita misura meno della meta dell'altezza."),
 ("Obiettivo pieno su tutti gli indici di forma", 69.0, 84, 96, "Rientro sotto il cut-off di 0,52 per WHtR e sotto 1,25 per il Conicity Index."),
]
r2 = r + 2
t0 = r2
for i, (nome, w, wc, hip, comm) in enumerate(SC2, 1):
    ws.cell(row=r2, column=1, value=i).alignment = Alignment(horizontal="center")
    ws.cell(row=r2, column=2, value=nome).font = Font(name=FONT, size=10, bold=(i == 1))
    for col, val in ((3, w), (4, wc), (5, hip)):
        c = ws.cell(row=r2, column=col, value=val)
        c.font = Font(name=FONT, size=10, bold=True, color=BLUE)
        c.fill = PatternFill("solid", fgColor=YEL)
    ws.cell(row=r2, column=6, value=f"=C{r2}/({I_H}/100)^2").number_format = "0.00"
    ws.cell(row=r2, column=7, value=f"=D{r2}/{I_H}").number_format = "0.000"
    ws.cell(row=r2, column=8, value=f"=(D{r2}/100)/(0.109*SQRT(C{r2}/({I_H}/100)))").number_format = "0.00"
    ws.cell(row=r2, column=9, value=f"=(D{r2}-65)*({I_TG}/88.57)").number_format = "0.0"
    ws.cell(row=r2, column=10, value=f"=(D{r2}/{I_H})*(({I_TG}/88.57)/({I_HDL}/38.67))").number_format = "0.00"
    ws.cell(row=r2, column=11, value=f"=(D{r2}/(39.68+1.88*F{r2}))*(({I_TG}/88.57)/1.03)*(1.31/({I_HDL}/38.67))").number_format = "0.00"
    ws.cell(row=r2, column=12, value=comm).font = Font(name=FONT, size=9)
    for cc in range(1, 13):
        cell = ws.cell(row=r2, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        cell.alignment = Alignment(horizontal="center" if 3 <= cc <= 11 or cc == 1 else "left",
                                   vertical="center", wrap_text=(cc in (2, 12)))
    ws.row_dimensions[r2].height = rowh(comm, 56, 28, 52)
    r2 += 1
t1 = r2 - 1

nr = r2 + 1
ws.cell(row=nr, column=2, value="Nota: i cut-off da rientrare sono WHtR 0,52 - Conicity 1,25 - LAP 26,7 - CMI 0,39 - VAI 1,92. "
                                "Gli indici ibridi migliorano anche solo riducendo la vita, perche la vita e il loro primo fattore.").font = Font(name=FONT, size=9, italic=True, color="595959")
ws.merge_cells(start_row=nr, start_column=2, end_row=nr + 1, end_column=12)
ws.cell(row=nr, column=2).alignment = Alignment(wrap_text=True, vertical="top")

# =====================================================================
# 08 CONTROLLI
# =====================================================================
ws = wb.create_sheet("08 Controlli")
setw(ws, {"A": 5, "B": 11, "C": 38, "D": 17, "E": 19, "F": 20, "G": 74})
title(ws, "08 - Controlli di riproducibilita",
      "Ogni valore pubblicato dalle fonti e stato ricalcolato dai dati grezzi. Qui e registrato cosa torna e cosa no.")
hdr(ws, 4, ["#", "Fonte", "Elemento verificato", "Valore pubblicato", "Valore ricalcolato", "Esito", "Impatto sul giudizio"])

CTRL = [
 (A2, "BMI da peso e altezza", "24,09", "24,44", "Scostamento",
  "74 kg su 174 cm danno 24,44. Per ottenere 24,09 servirebbero 72,9 kg a 174 cm, oppure 175,3 cm a 74 kg. Non e un arrotondamento: uno dei tre numeri pubblicati e incoerente con gli altri due. Il BMI entra nel VAI, nell'ABSI e nella correzione del polpaccio, quindi l'errore si propaga su tre indici."),
 (A2, "BMI usato internamente dal referto", "-", "24,09", "Dedotto",
  "Il VAI pubblicato (2,29) si riproduce solo con BMI 24,09; con 24,44 si otterrebbe 2,27. Il referto usa dunque internamente il BMI dichiarato, non quello ricavabile da peso e altezza. Il modello mantiene entrambi i valori senza sceglierne uno."),
 (A2, "Abdominal Volume Index", "11,89", "18,82", "Non riproducibile",
  "Scostamento del 58%. La formula standard (2 x vita^2 + 0,7 x (vita-fianchi)^2)/1000 non restituisce il valore pubblicato, e non e stata trovata alcuna variante che lo faccia. L'esito qualitativo non cambia: l'indice resta dentro il riferimento con entrambi i valori."),
 (A2, "Percentuale di grasso corporeo", "23%", "24,6% (Deurenberg)", "Non riproducibile",
  "La formula non e dichiarata. Deurenberg e la candidata piu probabile ma restituisce 24,6% con il BMI ricalcolato e 24,2% con quello del referto. Scostamento di circa 1,2 punti. Poco rilevante: e comunque una stima derivata dal BMI."),
 (A2, "Circonferenza polpaccio corretta per BMI", "33,5 cm", "34,0 cm", "Non riproducibile",
  "IMPATTO ALTO. La correzione categoriale del consenso EWGSOP2 non prevede aggiustamenti per BMI fra 18,5 e 24,9, quindi il valore resterebbe 34,0 cm, cioe esattamente sulla soglia invece che sotto. La diagnosi di sarcopenia moderata dipende interamente da una correzione di 5 millimetri la cui formula non e pubblicata."),
 (A2, "Conteggio degli indici oltre soglia", "6 su 12", "9 su 12", "Scostamento",
  "Applicando i cut-off riportati dalle tabelle della fonte stessa risultano oltre soglia: vita/altezza, vita/fianchi, conicita, BAI, LAP, CMI, VAI, ABSI e polpaccio. Restano dentro solo percentuale di grasso, BRI e AVI. La direzione della conclusione non cambia, anzi si rafforza."),
 (A2, "ABSI z-score", "1,52", "non ricalcolabile", "Non verificabile",
  "Il valore grezzo di ABSI si ricalcola (0,0873), ma lo z-score richiede le tabelle NHANES per eta e sesso, che la fonte non riporta. Nel foglio 02 il valore viene inserito a mano e non e verificato."),
 (A1, "Numero di indici cardiologici", "16", "15 tabulati", "Scostamento",
  "La fonte dichiara sedici indici ma le tre tabelle ne contengono quindici. Il modello include la superficie corporea come sedicesima riga informativa, essendo l'unico altro valore numerico del referto citato nel testo."),
 (A1, "Pressione arteriosa media", "101 mmHg", "98,7 (classica) / 101,6 (DBP+0,4xPP)", "Formula ricostruita",
  "La formula classica non riproduce il valore pubblicato. Rilevante perche la MAP e il denominatore del Modified Shock Index: con la formula classica il MSI passerebbe da 0,81 a 0,83."),
 (A1, "Left Cardiac Work Index", "627 mW/m2", "639 mW/m2", "Scostamento minore",
  "Circa il 2%, interamente dovuto agli arrotondamenti a monte: usando il Cardiac Index arrotondato a 2,8 si ottengono 628. Nessun impatto sul giudizio."),
 (A1, "Resistenze vascolari sistemiche", "2,7 kU*m2", "2,70 con CVP assunta a 5 mmHg", "Riproducibile con assunzione",
  "La pressione venosa centrale non e pubblicata. Con CVP 5 mmHg il calcolo coincide, con CVP 0 si otterrebbe 2,8, che spingerebbe l'indice sopra il limite superiore di 2,8. L'assunzione e quindi tutt'altro che neutra."),
 (A1, "Indice di rigidita arteriosa", "0,44", "0,436 (PP/MAP)", "Formula ricostruita",
  "Riproduce il valore pubblicato. Il rapporto alternativo PP/SBP darebbe 0,34."),
 (A1, "Zone di frequenza cardiaca", "87 / 105 / 122 / 140 / 157", "identiche con troncamento", "Riproducibile",
  "Le soglie si riproducono esattamente troncando all'intero le percentuali della FC massima, non arrotondando."),
 (A3, "Coerenza interna del profilo marziale", "-", "tutte le relazioni tornano", "Riproducibile",
  "Il pannello marziale e l'unico dei tre a essere internamente coerente al cento per cento: TIBC come doppio della transferrina, UIBC come differenza, saturazione come rapporto, saturazione aggiustata sul ferro corretto e depositi come ferritina per 4,5 riproducono tutti esattamente i valori pubblicati."),
 (A3, "Fattori di aggiustamento BRINDA", "-", "1,2574 sul ferro; 0,7502 sulla ferritina", "Dedotto",
  "I fattori si ricavano dal rapporto fra valori aggiustati e grezzi, ma non si verificano: la correzione BRINDA dipende dai reattanti di fase acuta e questo articolo non pubblica alcun marcatore di infiammazione, pur costruendoci sopra tutta l'argomentazione. Quello sulla ferritina e una riduzione del 25% esatto."),
 (A3, "Transferrina / log(ferritina)", "1,63", "1,677 (da 1,59 a 1,70 secondo i pesi molecolari)", "Sensibile alle unita",
  "La forma della formula e certa, il valore dipende dai pesi molecolari usati per convertire dalle unita SI a quelle convenzionali, che la fonte non dichiara. Il valore pubblicato cade dentro la banda dei valori plausibili. Con entrambi l'indice resta sotto il cut-off, ma a ridosso."),
 (A3, "Depositi di ferro: quale ferritina", "724 nM", "724 da ferritina grezza; 543 da ferritina aggiustata", "Scelta non dichiarata",
  "Il referto calcola i depositi sulla ferritina NON aggiustata, e cosi facendo li colloca sopra la finestra ottimale. Calcolandoli sulla ferritina aggiustata, che lo stesso referto pubblica due righe sopra, cadrebbero dentro la finestra. Le due righe raccontano storie diverse e la scelta non e motivata."),
 (A3, "Conteggio dei valori fuori finestra ottimale", "5 su 11", "6 su 11", "Scostamento",
  "Fuori dalle finestre ottimali risultano TIBC, UIBC, saturazione, saturazione aggiustata, ferritina aggiustata e depositi. L'articolo ne dichiara cinque."),
 (A3, "Coerenza fra testo e tabella", "'nessuna riga in rosso'", "saturazione 17,3 sotto il riferimento 20-48", "Contraddizione interna",
  "L'articolo apre dicendo che nessun valore e in rosso, e poche righe dopo afferma che la saturazione e l'unico valore sotto il riferimento anche senza aggiustamenti. Le due frasi non possono essere vere insieme: rispetto all'intervallo pubblicato in tabella, la saturazione e fuori."),
 (A4, "Ricostruzione dell'emocromo per inversione", "non pubblicato", "P 272; L 1,741; N 5,868; M 0,662", "Riproducibile",
  "L'articolo non pubblica i conteggi assoluti, ma il sistema di sei indici in quattro incognite e sovradeterminato e ammette soluzione. La verifica incrociata restituisce SII, SIRI, AISI e MLR esatti al centesimo, NLR entro lo 0,8% e PLR entro l'1,1%: la coerenza interna della sezione e alta."),
 (A4, "Emoglobina e RDW", "non pubblicati", "RDW 14,5; Hb 13,9 (scelti)", "Sottodeterminato",
  "Il rapporto RDW/Piastrine e pubblicato con due sole cifre decimali (0,05), il che lascia l'RDW libero fra 12,2 e 15,0 e con esso l'emoglobina fra 11,8 e 14,4. Sono stati scelti i valori coerenti con l'anemia lieve descritta dalla serie. E l'unica parte del quarto articolo che resta un'ipotesi."),
 (A4, "Conteggio degli indici fuori riferimento", "10 su 16", "9 fuori + 1 sul limite", "Coerente",
  "Il rapporto monociti/linfociti vale esattamente 0,38 contro un limite superiore di 0,38. Contandolo, i dieci dichiarati tornano esattamente. E l'unico dei quattro articoli il cui conteggio non presenta discrepanze."),
 (A4, "Rapporto monociti/HDL: quali unita", "16,1 per mille", "16,13 con HDL in mg/dL", "Ricostruita",
  "La forma per mille con l'HDL in mg/dL riproduce il valore pubblicato. La forma piu diffusa in letteratura, con HDL in mmol/L, darebbe 0,62: un numero che non ha alcun rapporto con il cut-off di 6 indicato nella stessa riga."),
 ("Art.3 + Art.4", "Aggiustamento per infiammazione: verifica incrociata", "ferritina ridotta del 25%", "PCR 0,3 mg/dL, cioe 3 mg/L", "Tensione fra le fonti",
  "PUNTO PIU DELICATO DEL WORKBOOK. L'articolo 3 applica alla ferritina una riduzione del 25% per infiammazione; l'articolo 4 pubblica una PCR di 3 mg/L, sotto la soglia convenzionale di 5 mg/L oltre la quale la correzione BRINDA entra in gioco. Con la sola PCR quella riduzione appare piu ampia del dovuto. L'alfa-1-glicoproteina acida, l'altro reattante su cui BRINDA si basa, non e pubblicata da nessuna delle due fonti, quindi la correzione resta possibile ma non verificabile. Sul piano fisiopatologico i dieci indici derivati alterati la sostengono; sul piano del calcolo, no."),
 ("Entrambe", "Tutti gli altri indici", "-", "coincidono", "Riproducibile",
  "Riproducono il valore pubblicato entro l'arrotondamento: pressione di pulsazione, riserva, Modified Shock Index, prodotto cardiovascolare, portata, Cardiac Index, morfotipo, pesi ideali di Lorenz e Creff, vita/altezza, vita/fianchi, Conicity Index, BAI, BRI, LAP, CMI e VAI."),
]
r = 5
for i, (fonte, elem, pub, calc, esito, imp) in enumerate(CTRL, 1):
    ws.cell(row=r, column=1, value=i).alignment = Alignment(horizontal="center", vertical="top")
    ws.cell(row=r, column=2, value=fonte.split(" - ")[0]).font = Font(name=FONT, size=9, color="595959")
    ws.cell(row=r, column=3, value=elem).font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=4, value=pub).font = Font(name=FONT, size=10, color=BLUE)
    ws.cell(row=r, column=5, value=calc).font = Font(name=FONT, size=10, bold=True)
    bad = esito in ("Non riproducibile", "Scostamento", "Non verificabile")
    ws.cell(row=r, column=6, value=esito).font = Font(name=FONT, size=10, bold=True,
                                                     color="C00000" if bad else "008000")
    ws.cell(row=r, column=7, value=imp).font = Font(name=FONT, size=9)
    ws.row_dimensions[r].height = rowh(imp, 90, 30, 84)
    for cc in range(1, 8):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        cell.alignment = Alignment(wrap_text=True, vertical="top",
                                   horizontal="center" if cc in (1, 2, 4, 5, 6) else "left")
        if bad:
            cell.fill = ERR_FILL
    r += 1

sr = r + 1
ws.cell(row=sr, column=3, value="Sintesi").font = Font(name=FONT, size=10, bold=True)
for k, (lab, val) in enumerate([
        ("Elementi verificati", f"=COUNTA(C5:C{r-1})"),
        ("Riproducibili senza riserve", f'=COUNTIF(F5:F{r-1},"Riproducibile")'),
        ("Ricostruiti o dipendenti da assunzioni dell'analista",
         f'=COUNTIF(F5:F{r-1},"Formula ricostruita")+COUNTIF(F5:F{r-1},"Riproducibile con assunzione")'
         f'+COUNTIF(F5:F{r-1},"Dedotto")+COUNTIF(F5:F{r-1},"Sensibile alle unita")'),
        ("Scostamenti, valori non riproducibili o incongruenze", f"=E{sr+1}-E{sr+2}-E{sr+3}")]):
    ws.cell(row=sr + 1 + k, column=3, value=lab).font = Font(name=FONT, size=10)
    c = ws.cell(row=sr + 1 + k, column=5, value=val)
    c.font = Font(name=FONT, size=11, bold=True, color="C00000" if k == 3 else BLACK)
    c.alignment = Alignment(horizontal="center")
    for cc in range(3, 7):
        ws.cell(row=sr + 1 + k, column=cc).border = BORDER

cr2 = sr + 6
ws.cell(row=cr2, column=3, value="Come leggere questo foglio").font = Font(name=FONT, size=10, bold=True, color="C00000")
c = ws.cell(row=cr2 + 1, column=3, value=(
    "Nessuno di questi scostamenti ribalta le conclusioni delle fonti: gli indici che segnalano rischio lo segnalano "
    "con entrambi i valori, e il conteggio degli indici oltre soglia risulta piu alto del dichiarato, non piu basso. "
    "Due elementi meritano pero attenzione. Il primo e la diagnosi di sarcopenia, che poggia su una correzione di "
    "cinque millimetri la cui formula non e pubblicata e che con il criterio del consenso europeo non verrebbe "
    "applicata. Il secondo e il BMI, che non discende dal peso e dall'altezza dichiarati e che alimenta tre indici "
    "derivati. Entrambi si risolvono chiedendo alla fonte i dati grezzi o la formula usata."))
c.font = Font(name=FONT, size=9, italic=True)
c.alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=cr2 + 1, start_column=3, end_row=cr2 + 4, end_column=7)
for cc in range(3, 8):
    for rr in range(cr2 + 1, cr2 + 5):
        ws.cell(row=rr, column=cc).fill = WARN_FILL
        ws.cell(row=rr, column=cc).border = BORDER

# =====================================================================
# 09 CONVERSIONI
# =====================================================================
ws = wb.create_sheet("09 Conversioni")
setw(ws, {"A": 5, "B": 36, "C": 14, "D": 9, "E": 10, "F": 16, "G": 12, "H": 60})
title(ws, "09 - Profilo marziale in unita convenzionali",
      "Il referto usa unita SI, che quasi nessun laboratorio italiano stampa. Qui gli stessi valori nelle unita di un referto ordinario.")
hdr(ws, 4, ["#", "Analita", "Valore SI", "Unita SI", "Fattore", "Valore convenzionale", "Unita", "Nota"])

CONV = [
 ("Sideremia", C(32), "uM", K_FE_UGDL, "ug/dL", "0.0",
  "Il valore convenzionale corrispondente e appena sopra il limite inferiore tipico per l'uomo adulto."),
 ("Sideremia aggiustata", C(33), "uM", K_FE_UGDL, "ug/dL", "0.0", ""),
 ("Transferrina", C(34), "uM", MW_TRF / 1e4, "mg/dL", "0.0",
  "Convertita con peso molecolare 79.570 g/mol."),
 ("TIBC", C(35), "uM", K_FE_UGDL, "ug/dL", "0.0",
  "Il valore convenzionale e sopra l'intervallo tipico: molti posti liberi sui camion."),
 ("UIBC", C(36), "uM", K_FE_UGDL, "ug/dL", "0.0", ""),
 ("Ferritina", C(39), "pM", K_FER_UGL, "ug/L (ng/mL)", "0.0",
  "Convertita con peso molecolare 450.000 g/mol. Nelle unita di un referto ordinario e un valore che nessuno segnalerebbe."),
 ("Ferritina aggiustata", C(40), "pM", K_FER_UGL, "ug/L (ng/mL)", "0.0",
  "Sotto la soglia di 30 ug/L nessuno discuterebbe la carenza; qui il valore resta sopra, ed e per questo che il caso e difficile."),
 ("Soglia ottimale di ferritina aggiustata", "=283.1", "pM", K_FER_UGL, "ug/L (ng/mL)", "0.0",
  "Il riferimento ottimale del referto tradotto in unita convenzionali."),
]
r = 5
for i, (nome, sival, siu, k, cu, fmt, nota) in enumerate(CONV, 1):
    ws.cell(row=r, column=1, value=i).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=2, value=nome).font = Font(name=FONT, size=10, bold=True)
    c = ws.cell(row=r, column=3, value=f"={sival}" if not str(sival).startswith("=") else sival)
    c.font = Font(name=FONT, size=10, color=GREEN)
    c.number_format = "0.0"
    ws.cell(row=r, column=4, value=siu)
    ws.cell(row=r, column=5, value=k).number_format = "0.0000"
    cv = ws.cell(row=r, column=6, value=f"=C{r}*E{r}")
    cv.font = Font(name=FONT, size=11, bold=True)
    cv.number_format = fmt
    ws.cell(row=r, column=7, value=cu)
    ws.cell(row=r, column=8, value=nota).font = Font(name=FONT, size=9, italic=True, color="595959")
    for cc in range(1, 9):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        cell.alignment = Alignment(horizontal="center" if cc in (1, 3, 4, 5, 6, 7) else "left",
                                   vertical="center", wrap_text=(cc == 8))
    ws.row_dimensions[r].height = rowh(nota, 66, 26, 46)
    r += 1

nr3 = r + 1
ws.cell(row=nr3, column=2, value="Perche questo foglio esiste").font = Font(name=FONT, size=10, bold=True)
c = ws.cell(row=nr3 + 1, column=2, value=(
    "I fattori di conversione della ferritina e della transferrina dipendono dal peso molecolare, che non e una "
    "costante universale: la ferritina e riportata in letteratura fra 440.000 e 474.000 g/mol a seconda della "
    "composizione in catene leggere e pesanti. E la ragione per cui il rapporto transferrina/log(ferritina) del "
    "foglio 02 non coincide al centesimo con quello pubblicato. Per sideremia, TIBC e UIBC il fattore 5,585 e "
    "invece esatto, essendo il peso atomico del ferro diviso dieci."))
c.font = Font(name=FONT, size=9, italic=True)
c.alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=nr3 + 1, start_column=2, end_row=nr3 + 3, end_column=8)
for cc in range(2, 9):
    for rr in range(nr3 + 1, nr3 + 4):
        ws.cell(row=rr, column=cc).fill = WARN_FILL
        ws.cell(row=rr, column=cc).border = BORDER

# =====================================================================
for sheet in wb.worksheets:
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None and (cell.font is None or cell.font.name != FONT):
                f = cell.font
                cell.font = Font(name=FONT, size=f.size or 10, bold=f.bold, italic=f.italic,
                                 color=f.color, underline=f.underline)
    sheet.sheet_view.showGridLines = False

wb.save(OUT)
print("scritto:", OUT, "| righe indici:", R0, "-", LAST)
