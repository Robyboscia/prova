#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera il workbook operativo per l'inserimento dei dati di un paziente.

Differenze rispetto al workbook di verifica degli articoli:
  - formule e soglie distinte per uomo e donna
  - nessun valore del caso dimostrativo cablato dentro
  - gli indici che non hanno tutti i dati necessari restituiscono "n.d."
    e restano fuori dal punteggio
  - punteggio globale unico, corretto per la ridondanza fra indici che
    condividono la stessa misura di partenza
"""

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule

OUT = "/home/user/prova/Referto_Indici_Paziente.xlsx"
FONT = "Arial"
BLUE, BLACK, GREEN, YEL = "0000FF", "000000", "008000", "FFFF00"
HDR = PatternFill("solid", fgColor="1F3864")
SUB = PatternFill("solid", fgColor="D9E2F3")
WARN = PatternFill("solid", fgColor="FFF2CC")
BAD = PatternFill("solid", fgColor="F8CBAD")
OPT = PatternFill("solid", fgColor="EDEDED")
thin = Side(style="thin", color="BFBFBF")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)


def hdr(ws, row, labels, h=30, start=1):
    for i, lab in enumerate(labels):
        c = ws.cell(row=row, column=start + i, value=lab)
        c.font = Font(name=FONT, size=10, bold=True, color="FFFFFF")
        c.fill = HDR
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = BORDER
    ws.row_dimensions[row].height = h


def title(ws, t, sub=None):
    ws.cell(row=1, column=1, value=t).font = Font(name=FONT, size=14, bold=True, color="1F3864")
    if sub:
        ws.cell(row=2, column=1, value=sub).font = Font(name=FONT, size=9, italic=True, color="595959")


def setw(ws, w):
    for k, v in w.items():
        ws.column_dimensions[k].width = v


def rowh(t, width=80, lo=26, hi=76):
    return max(lo, min(hi, 13 * (len(t) // width + 1)))


wb = Workbook()

# =====================================================================
# 01 PAZIENTE
# =====================================================================
ws = wb.create_sheet("01 Paziente")
setw(ws, {"A": 44, "B": 15, "C": 13, "D": 11, "E": 70})
title(ws, "01 - Dati del paziente",
      "Compila le celle gialle. Le celle grigie sono facoltative: se le lasci vuote gli indici che le usano restituiscono 'n.d.' e non entrano nel punteggio.")
hdr(ws, 4, ["Dato", "Valore", "Unita", "Obbl.", "Nota"])

INP = [
    ("IDENTIFICAZIONE", None, None, None, None),
    ("Identificativo / codice paziente", "", "", "si", "Non inserire dati identificativi diretti se il file viene condiviso"),
    ("Data della valutazione", "", "gg/mm/aaaa", "si", "Serve al foglio 05 per il confronto nel tempo"),
    ("Sesso", "M", "M / F", "si", "DETERMINANTE: cambia formule e soglie di 14 indici. Menu a tendina"),
    ("Eta", "", "anni", "si", ""),
    ("", None, None, None, None),
    ("ANTROPOMETRIA (metro da sarta e bilancia)", None, None, None, None),
    ("Peso", "", "kg", "si", ""),
    ("Altezza", "", "cm", "si", ""),
    ("Circonferenza vita", "", "cm", "si", "A fine espirazione, a meta fra ultima costa e cresta iliaca. E la misura piu importante del foglio: da sola alimenta sei indici"),
    ("Circonferenza fianchi", "", "cm", "no", "Serve a vita/fianchi, BAI e volume addominale"),
    ("Circonferenza polpaccio", "", "cm", "no", "Surrogato della massa muscolare. Soglia 34 cm nell'uomo, 33 nella donna"),
    ("Circonferenza polso", "", "cm", "no", "Morfotipo. E l'unica misura che non cambia con dieta o allenamento"),
    ("", None, None, None, None),
    ("PRESSIONE E CUORE", None, None, None, None),
    ("Pressione sistolica", "", "mmHg", "no", ""),
    ("Pressione diastolica", "", "mmHg", "no", ""),
    ("Frequenza cardiaca a riposo", "", "bpm", "no", "Da misurare seduto, dopo cinque minuti di riposo"),
    ("Gittata sistolica", "", "mL", "no", "Solo se disponibile da ecocardiogramma o bioimpedenza cardiaca"),
    ("", None, None, None, None),
    ("LIPIDI E METABOLISMO", None, None, None, None),
    ("Colesterolo HDL", "", "mg/dL", "no", "Entra in quattro indici: e il denominatore piu ricorrente del foglio"),
    ("Trigliceridi", "", "mg/dL", "no", ""),
    ("Uricemia", "", "mg/dL", "no", ""),
    ("Omocisteina", "", "uM", "no", ""),
    ("", None, None, None, None),
    ("ASSETTO MARZIALE", None, None, None, None),
    ("Sideremia", "", "ug/dL", "no", "In unita convenzionali. Il prelievo va fatto al mattino a digiuno: il ferro ha un ritmo circadiano marcato"),
    ("Transferrina", "", "mg/dL", "no", ""),
    ("Ferritina", "", "ug/L", "no", "Non interpretabile da sola se c'e infiammazione: il foglio lo segnala"),
    ("", None, None, None, None),
    ("INFIAMMAZIONE ED EMOCROMO", None, None, None, None),
    ("Proteina C reattiva (PCR)", "", "mg/dL", "no", "Serve anche a stabilire se la ferritina e leggibile"),
    ("VES a 1 ora", "", "mm", "no", ""),
    ("Emoglobina", "", "g/dL", "no", ""),
    ("RDW", "", "%", "no", ""),
    ("Piastrine", "", "10^9/L", "no", ""),
    ("Volume piastrinico medio (MPV)", "", "fL", "no", ""),
    ("Neutrofili (valore assoluto)", "", "10^9/L", "no", "Valore ASSOLUTO, non percentuale"),
    ("Linfociti (valore assoluto)", "", "10^9/L", "no", "Valore ASSOLUTO. Da solo alimenta sette indici"),
    ("Monociti (valore assoluto)", "", "10^9/L", "no", "Valore ASSOLUTO"),
    ("", None, None, None, None),
    ("FACOLTATIVI E ANAMNESTICI", None, None, None, None),
    ("ABSI z-score", "", "z", "no", "Da inserire solo se il laboratorio lo fornisce: richiede le tabelle di riferimento per eta e sesso. Il foglio calcola comunque l'ABSI grezzo"),
    ("Sigarette al giorno", "", "n/die", "no", "0 se non fumatore"),
    ("Anni di abitudine tabagica", "", "anni", "no", ""),
    ("Sedentarieta (1 = si, 0 = no)", "", "0/1", "no", "Assenza di esercizio strutturato"),
    ("Russamento o pause respiratorie riferite (1/0)", "", "0/1", "no", "Domanda da porre sempre: non compare in nessun referto"),
    ("", None, None, None, None),
    ("CALCOLATI AUTOMATICAMENTE - non compilare", None, None, None, None),
]
r = 5
IR = {}
for nome, val, unit, obbl, nota in INP:
    if nome == "":
        r += 1
        continue
    if val is None:
        c = ws.cell(row=r, column=1, value=nome)
        c.font = Font(name=FONT, size=10, bold=True, color="1F3864")
        for cc in range(1, 6):
            ws.cell(row=r, column=cc).fill = SUB
            ws.cell(row=r, column=cc).border = BORDER
        r += 1
        continue
    ws.cell(row=r, column=1, value=nome).font = Font(name=FONT, size=10)
    vc = ws.cell(row=r, column=2, value=val if val != "" else None)
    vc.font = Font(name=FONT, size=11, bold=True, color=BLUE)
    vc.fill = PatternFill("solid", fgColor=YEL) if obbl == "si" else OPT
    vc.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=3, value=unit).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=4, value=obbl).alignment = Alignment(horizontal="center")
    nc = ws.cell(row=r, column=5, value=nota)
    nc.font = Font(name=FONT, size=9, italic=True, color="595959")
    nc.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = rowh(nota, 86, 16, 40)
    for cc in range(1, 6):
        ws.cell(row=r, column=cc).border = BORDER
        if ws.cell(row=r, column=cc).font.name != FONT:
            ws.cell(row=r, column=cc).font = Font(name=FONT, size=10)
    IR[nome] = r
    r += 1

P = lambda k: f"'01 Paziente'!$B${IR[k]}"
SEX = P("Sesso")
AGE = P("Eta")
W, H = P("Peso"), P("Altezza")
WC, HIP = P("Circonferenza vita"), P("Circonferenza fianchi")
CALF, WRIST = P("Circonferenza polpaccio"), P("Circonferenza polso")
SBP, DBP = P("Pressione sistolica"), P("Pressione diastolica")
HR, SV = P("Frequenza cardiaca a riposo"), P("Gittata sistolica")
HDL, TG = P("Colesterolo HDL"), P("Trigliceridi")
URIC, HCY = P("Uricemia"), P("Omocisteina")
FE, TRF, FER = P("Sideremia"), P("Transferrina"), P("Ferritina")
PCR, VES = P("Proteina C reattiva (PCR)"), P("VES a 1 ora")
HB, RDW = P("Emoglobina"), P("RDW")
PLT, MPV = P("Piastrine"), P("Volume piastrinico medio (MPV)")
NEU, LYM, MON = P("Neutrofili (valore assoluto)"), P("Linfociti (valore assoluto)"), P("Monociti (valore assoluto)")
ABSIZ = P("ABSI z-score")
CIG, YRS = P("Sigarette al giorno"), P("Anni di abitudine tabagica")
SED, SNORE = P("Sedentarieta (1 = si, 0 = no)"), P("Russamento o pause respiratorie riferite (1/0)")

# blocco calcolati
CALC = [
    ("Indice di massa corporea (BMI)", f"=IF(OR({W}=\"\",{H}=\"\"),\"\",{W}/({H}/100)^2)", "kg/m2",
     "Calcolato da peso e altezza: nessun BMI va inserito a mano"),
    ("Superficie corporea (Mosteller)", f"=IF(OR({W}=\"\",{H}=\"\"),\"\",SQRT({H}*{W}/3600))", "m2",
     "Formula di Mosteller. Sostituisce l'inserimento manuale"),
    ("Frequenza cardiaca massima stimata", f'=IF({AGE}="","",IF({SEX}="M",208-0.7*{AGE},206-0.88*{AGE}))', "bpm",
     "Tanaka nell'uomo, Gulati nella donna. E una STIMA di popolazione: l'intervallo atteso e il 90-110% del valore, non il valore secco"),
    ("Sesso in forma numerica", f'=IF({SEX}="M",1,0)', "1/0", "Usato dalle formule che hanno un termine per il sesso"),
    ("Ferritina interpretabile", f'=IF({PCR}="","dato mancante",IF({PCR}>0.5,"NO - infiammazione in atto","si"))', "",
     "Con PCR sopra 0,5 mg/dL la ferritina misura anche l'infiammazione e non e leggibile come riserva di ferro"),
]
for nome, f, unit, nota in CALC:
    ws.cell(row=r, column=1, value=nome).font = Font(name=FONT, size=10)
    c = ws.cell(row=r, column=2, value=f)
    c.font = Font(name=FONT, size=11, bold=True, color=BLACK)
    c.number_format = "0.00"
    c.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=3, value=unit).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=4, value="auto").alignment = Alignment(horizontal="center")
    nc = ws.cell(row=r, column=5, value=nota)
    nc.font = Font(name=FONT, size=9, italic=True, color="595959")
    nc.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = rowh(nota, 86, 16, 40)
    for cc in range(1, 6):
        ws.cell(row=r, column=cc).border = BORDER
        if ws.cell(row=r, column=cc).font.name != FONT:
            ws.cell(row=r, column=cc).font = Font(name=FONT, size=10)
    IR[nome] = r
    r += 1

BMI = P("Indice di massa corporea (BMI)")
BSA = P("Superficie corporea (Mosteller)")
HRMAX = P("Frequenza cardiaca massima stimata")
SEXN = P("Sesso in forma numerica")

dv = DataValidation(type="list", formula1='"M,F"', allow_blank=False)
ws.add_data_validation(dv)
dv.add(ws.cell(row=IR["Sesso"], column=2))

# =====================================================================
# 02 RIFERIMENTI  (soglie per sesso, modificabili)
# =====================================================================
ws = wb.create_sheet("02 Riferimenti")
setw(ws, {"A": 5, "B": 40, "C": 11, "D": 11, "E": 11, "F": 11, "G": 12, "H": 7, "I": 16, "J": 52})
title(ws, "02 - Soglie di riferimento per sesso",
      "Tutte le soglie sono modificabili: allineale al tuo laboratorio. Il foglio 03 le legge da qui in base al sesso indicato nel foglio 01.")
hdr(ws, 4, ["#", "Indice", "Min uomo", "Max uomo", "Min donna", "Max donna",
            "Direzione critica", "Peso", "Gruppo (ridondanza)", "Origine della soglia"])

# (nome, unita, formula, minM, maxM, minF, maxF, direzione, peso, gruppo, fmt, origine, nota)
IDX = [
 # ---- antropometria
 ("Indice di massa corporea (BMI)", "kg/m2", f"={BMI}", 18.5, 24.9, 18.5, 24.9, "Informativo", 0, "Corporatura", "0.00",
  "OMS", "Peso 0 di proposito: e l'indice che assolve i corpi in disordine. Serve come contesto, non come giudizio."),
 ("Rapporto vita / altezza", "-", f'=IF(OR({WC}="",{H}=""),"n.d.",{WC}/{H})', 0.30, 0.50, 0.30, 0.50, "Alto", 3, "Vita", "0.00",
  "Ashwell, PMID 22106927", "La regola aurea vale per entrambi i sessi: la vita deve misurare meno della meta dell'altezza. E il primo indice da guardare."),
 ("Rapporto vita / fianchi", "-", f'=IF(OR({WC}="",{HIP}=""),"n.d.",{WC}/{HIP})', 0.70, 0.90, 0.70, 0.85, "Alto", 2, "Vita", "0.00",
  "OMS", "Soglia diversa per sesso: 0,90 nell'uomo, 0,85 nella donna."),
 ("Conicity Index", "-", f'=IF(OR({WC}="",{W}="",{H}=""),"n.d.",({WC}/100)/(0.109*SQRT({W}/({H}/100))))', 1.10, 1.25, 1.05, 1.18, "Alto", 2, "Vita", "0.00",
  "Valdez", "Quanto il tronco si allontana dal cilindro. Soglia piu bassa nella donna."),
 ("Percentuale di grasso corporeo (stima)", "%", f'=IF(OR({BMI}="",{AGE}=""),"n.d.",1.2*{BMI}+0.23*{AGE}-10.8*{SEXN}-5.4)', 8, 25, 20, 36, "Alto", 1, "Corporatura", "0.0",
  "Deurenberg", "STIMA derivata dal BMI, con errore di alcuni punti percentuali. La formula ha un termine per il sesso e le soglie sono molto diverse fra uomo e donna."),
 ("Body Adiposity Index", "%", f'=IF(OR({HIP}="",{H}=""),"n.d.",{HIP}/(({H}/100)^1.5)-18)', 8, 24, 21, 35, "Alto", 1, "Fianchi", "0.0",
  "Bergman", "Peso ridotto a 1: sovrastima nei soggetti magri con fianchi normali. Non usa mai il peso: utile quando diverge dalla stima basata sul BMI. Fasce di normalita molto diverse per sesso."),
 ("Body Roundness Index", "-", f'=IF(OR({WC}="",{H}=""),"n.d.",364.2-365.5*SQRT(1-((({WC}/100)/(2*PI()))/(0.5*({H}/100)))^2))', 1.0, 4.71, 1.0, 4.45, "Alto", 1, "Vita", "0.00",
  "Thomas", "Soglia indicativa, piu bassa nella donna."),
 ("A Body Shape Index (grezzo)", "-", f'=IF(OR({WC}="",{BMI}="",{H}=""),"n.d.",({WC}/100)/(({BMI}^(2/3))*SQRT({H}/100)))', 0.070, 0.083, 0.068, 0.081, "Alto", 3, "Vita", "0.0000",
  "Krakauer, PMID 22815707", "Isola la forma corporea depurata da peso e BMI. Se il laboratorio fornisce lo z-score, inseriscilo nel foglio 01: e piu preciso di questa soglia grezza."),
 ("Circonferenza polpaccio", "cm", f'=IF({CALF}="","n.d.",{CALF})', 34, 45, 33, 43, "Basso", 3, "Muscolo", "0.0",
  "EWGSOP2, PMID 30312372", "Soglia 34 cm nell'uomo, 33 nella donna. Nessuna correzione per BMI: quella usata da alcuni referti non e pubblicata e sposta la diagnosi di sarcopenia."),
 ("Peso ideale (Lorenz)", "kg", f'=IF(OR({H}="",{SEX}=""),"n.d.",{H}-100-({H}-150)/IF({SEX}="M",4,2))', 40, 120, 40, 120, "Informativo", 0, "Corporatura", "0.0",
  "Lorenz", "Denominatore diverso per sesso. Indicativo: e una convenzione, non una misura."),
 ("Morfotipo (altezza / polso)", "-", f'=IF(OR({H}="",{WRIST}=""),"n.d.",{H}/{WRIST})', 9.6, 10.4, 10.1, 11.0, "Informativo", 0, "Corporatura", "0.00",
  "Convenzione antropometrica", "Fasce diverse per sesso. Dentro la fascia = normolineo; sotto = brevilineo; sopra = longilineo."),
 # ---- cuore e pressione
 ("Pressione sistolica", "mmHg", f'=IF({SBP}="","n.d.",{SBP})', 100, 129, 100, 129, "Alto", 1, "Pressione", "0",
  "ESC/ESH", ""),
 ("Pressione diastolica", "mmHg", f'=IF({DBP}="","n.d.",{DBP})', 60, 84, 60, 84, "Alto", 1, "Pressione", "0",
  "ESC/ESH", ""),
 ("Pressione arteriosa media", "mmHg", f'=IF(OR({SBP}="",{DBP}=""),"n.d.",({SBP}+2*{DBP})/3)', 70, 105, 70, 105, "Alto", 1, "Pressione", "0",
  "Formula classica", "Qui si usa la formula classica (SBP+2xDBP)/3. Alcuni referti usano DBP+0,4xPP, che da un valore piu alto: se confronti con un referto, verifica quale delle due ha usato."),
 ("Pressione di pulsazione", "mmHg", f'=IF(OR({SBP}="",{DBP}=""),"n.d.",{SBP}-{DBP})', 30, 55, 30, 55, "Alto", 2, "Pressione", "0",
  "PMID 33793325", "Marcatore di rigidita dei grossi vasi e di invecchiamento vascolare."),
 ("Frequenza cardiaca a riposo", "bpm", f'=IF({HR}="","n.d.",{HR})', 50, 76, 50, 80, "Alto", 3, "Frequenza", "0",
  "PMID 26598376", "Sopra 80 bpm il rischio cambia scala. Soglia leggermente piu alta nella donna."),
 ("Frequenza cardiaca di riserva", "bpm", f'=IF(OR({HR}="",{HRMAX}=""),"n.d.",{HRMAX}-{HR})', 70, 130, 70, 130, "Basso", 1, "Frequenza", "0",
  "PMID 17446799", "Attenzione all'artefatto: migliora abbassando il basale, senza che il cuore cambi."),
 ("Modified Shock Index", "-", f'=IF(OR({HR}="",{SBP}="",{DBP}=""),"n.d.",{HR}/(({SBP}+2*{DBP})/3))', 0.70, 0.99, 0.70, 0.99, "Alto", 1, "Frequenza", "0.00",
  "Letteratura", "Va calcolato sulla pressione MEDIA, non sulla sistolica."),
 ("Prodotto cardiovascolare", "bpm*mmHg", f'=IF(OR({HR}="",{SBP}=""),"n.d.",{HR}*{SBP})', 6000, 12000, 6000, 12000, "Alto", 2, "Frequenza", "#,##0",
  "PMID 38453019", "Surrogato del consumo miocardico di ossigeno. Tre mesi di esercizio supervisionato lo abbassano di circa il 19%."),
 ("Portata cardiaca", "L/min", f'=IF(OR({SV}="",{HR}=""),"n.d.",{SV}*{HR}/1000)', 3.5, 6.9, 3.0, 6.5, "Basso", 1, "Pompa", "0.0",
  "Fisiologia", "Richiede la gittata sistolica, spesso non disponibile."),
 ("Cardiac Index", "L/min/m2", f'=IF(OR({SV}="",{HR}="",{BSA}=""),"n.d.",{SV}*{HR}/1000/{BSA})', 2.2, 4.0, 2.2, 4.0, "Basso", 1, "Pompa", "0.00",
  "Fisiologia", ""),
 # ---- lipidi e metabolismo
 ("Colesterolo HDL", "mg/dL", f'=IF({HDL}="","n.d.",{HDL})', 40, 90, 50, 100, "Basso", 2, "HDL", "0",
  "ATP III", "Soglia 40 nell'uomo, 50 nella donna. Compare come denominatore in quattro indici di questo foglio."),
 ("Lipid Accumulation Product", "-", f'=IF(OR({WC}="",{TG}=""),"n.d.",({WC}-IF({SEX}="M",65,58))*({TG}/88.57))', 0, 26.7, 0, 22.4, "Alto", 3, "Vita", "0.0",
  "Kahn, PMID 16150143", "Costante di sottrazione diversa per sesso (65 cm uomo, 58 donna) e cut-off diverso. Riconosce il rischio cardiovascolare meglio del BMI."),
 ("Cardiometabolic Index", "-", f'=IF(OR({WC}="",{H}="",{TG}="",{HDL}=""),"n.d.",({WC}/{H})*(({TG}/88.57)/({HDL}/38.67)))', 0, 0.39, 0, 0.31, "Alto", 2, "Vita", "0.00",
  "Wakabayashi, PMID 25199852", "Da calcolare in mmol/L: con i mg/dL il numero non ha rapporto con il cut-off. Soglia piu bassa nella donna."),
 ("Visceral Adiposity Index", "-",
  f'=IF(OR({WC}="",{BMI}="",{TG}="",{HDL}=""),"n.d.",IF({SEX}="M",({WC}/(39.68+1.88*{BMI}))*(({TG}/88.57)/1.03)*(1.31/({HDL}/38.67)),({WC}/(36.58+1.89*{BMI}))*(({TG}/88.57)/0.81)*(1.52/({HDL}/38.67))))',
  0, 1.92, 0, 1.92, "Alto", 3, "Vita", "0.00",
  "Amato, PMID 20067971", "FORMULA COMPLETAMENTE DIVERSA PER SESSO: tutti e tre i gruppi di coefficienti cambiano. Misura la funzione del tessuto adiposo, non la sua quantita."),
 ("Uricemia", "mg/dL", f'=IF({URIC}="","n.d.",{URIC})', 3.5, 7.2, 2.6, 6.0, "Alto", 2, "Uricemia", "0.0",
  "Laboratorio", "Intervalli molto diversi per sesso."),
 ("Uricemia / HDL", "-", f'=IF(OR({URIC}="",{HDL}=""),"n.d.",{URIC}/{HDL}*100)', 0, 12.2, 0, 9.5, "Alto", 2, "Uricemia", "0.0",
  "Letteratura", "Soglia piu bassa nella donna, coerente con l'uricemia piu bassa."),
 ("Omocisteina", "uM", f'=IF({HCY}="","n.d.",{HCY})', 3, 15, 3, 15, "Alto", 2, "Omocisteina", "0.0",
  "Laboratorio", "Finestra desiderabile molto piu stretta del riferimento tollerato: 5,0-7,2 uM."),
 # ---- assetto marziale
 ("Saturazione della transferrina", "%", f'=IF(OR({FE}="",{TRF}=""),"n.d.",{FE}/({TRF}*1.404)*100)', 20, 48, 20, 48, "Basso", 3, "Ferro", "0.0",
  "PMID 27346617", "Il marcatore piu informativo dello stato marziale reale. TIBC ricavata dalla transferrina con il fattore stechiometrico esatto (mg/dL x 1,404 = ug/dL), non con la regola pratica dell'1,25 che sottostima la capacita legante di circa il 12%."),
 ("Ferritina", "ug/L", f'=IF({FER}="","n.d.",{FER})', 30, 400, 15, 150, "Basso", 2, "Ferro", "0",
  "Laboratorio", "INTERVALLI MOLTO DIVERSI PER SESSO. Da non leggere da sola se la PCR e alterata: il foglio 01 lo segnala."),
 ("Emoglobina", "g/dL", f'=IF({HB}="","n.d.",{HB})', 13.0, 17.5, 12.0, 15.5, "Basso", 3, "Emoglobina", "0.0",
  "OMS", "Soglia di anemia 13 g/dL nell'uomo, 12 nella donna."),
 ("Emoglobina / RDW", "-", f'=IF(OR({HB}="",{RDW}=""),"n.d.",{HB}/{RDW})', 1.0, 1.6, 0.9, 1.5, "Basso", 2, "Emoglobina", "0.00",
  "Letteratura", "Cala quando il midollo produce eritrociti di taglia irregolare."),
 # ---- infiammazione
 ("Proteina C reattiva", "mg/dL", f'=IF({PCR}="","n.d.",{PCR})', 0, 0.5, 0, 0.5, "Alto", 0, "Infiammazione classica", "0.00",
  "Laboratorio", "Peso 0: esame giusto per la domanda sbagliata. Intercetta l'infiammazione acuta, non quella cronica di basso grado. Serve pero a dire se la ferritina e leggibile."),
 ("VES a 1 ora", "mm", f'=IF({VES}="","n.d.",{VES})', 1, 15, 1, 20, "Alto", 0, "Infiammazione classica", "0",
  "Laboratorio", "Peso 0, stesso limite della PCR. Intervallo piu ampio nella donna."),
 ("Neutrofili / Linfociti", "-", f'=IF(OR({NEU}="",{LYM}=""),"n.d.",{NEU}/{LYM})', 0.73, 3.33, 0.73, 3.33, "Alto", 2, "Linfociti", "0.00",
  "Letteratura", "L'infiammazione cronica alza l'immunita innata e abbassa i linfociti: il rapporto lo vede anche quando i due valori assoluti sono normali."),
 ("Monociti / Linfociti", "-", f'=IF(OR({MON}="",{LYM}=""),"n.d.",{MON}/{LYM})', 0.12, 0.38, 0.12, 0.38, "Alto", 2, "Linfociti", "0.00",
  "Letteratura", ""),
 ("Piastrine / Linfociti", "-", f'=IF(OR({PLT}="",{LYM}=""),"n.d.",{PLT}/{LYM})', 63, 209, 63, 209, "Alto", 1, "Linfociti", "0",
  "Letteratura", ""),
 ("Systemic Inflammation Index", "-", f'=IF(OR({PLT}="",{NEU}="",{LYM}=""),"n.d.",{PLT}*{NEU}/{LYM})', 131, 901, 131, 901, "Alto", 2, "Linfociti", "0",
  "PMID 36769776", "Associato a mortalita totale e cardiovascolare su 42.875 adulti seguiti vent'anni."),
 ("Systemic Inflammation Response Index", "-", f'=IF(OR({NEU}="",{MON}="",{LYM}=""),"n.d.",{NEU}*{MON}/{LYM})', 0, 1.20, 0, 1.20, "Alto", 3, "Linfociti", "0.00",
  "PMID 36769776", "Soglia indicativa. Lo 0,68 usato da alcuni referti e un percentile basso e segnala anche emocromi del tutto normali: qui si usa un valore piu prudente. Taralo sul tuo laboratorio."),
 ("Aggregate Index of Systemic Inflammation", "-", f'=IF(OR({NEU}="",{MON}="",{PLT}="",{LYM}=""),"n.d.",{NEU}*{MON}*{PLT}/{LYM})', 0, 400, 0, 400, "Alto", 3, "Linfociti", "0.00",
  "Xiu, PMID 37265570", "Nel quartile piu alto il rischio di mortalita cardiovascolare quasi raddoppia (HR 1,91)."),
 ("Monociti / HDL", "per mille", f'=IF(OR({MON}="",{HDL}=""),"n.d.",{MON}/{HDL}*1000)', 0, 13, 0, 11, "Alto", 3, "HDL", "0.0",
  "Gembillo, PMID 34109496", "Mette in relazione chi accende l'infiammazione e chi dovrebbe spegnerla: le HDL ostacolano il reclutamento dei monociti nella parete arteriosa."),
 ("MPV / Linfociti", "-", f'=IF(OR({MPV}="",{LYM}=""),"n.d.",{MPV}/{LYM})', 0, 5.55, 0, 5.55, "Alto", 1, "Linfociti", "0.00",
  "Korniluk, PMID 31148950", ""),
 ("MPV / Piastrine", "-", f'=IF(OR({MPV}="",{PLT}=""),"n.d.",{MPV}/{PLT}*100)', 0, 6.0, 0, 6.0, "Alto", 1, "Piastrine", "0.00",
  "Korniluk, PMID 31148950", ""),
]

r = 5
REF0 = r
for i, (nome, unit, f, mnM, mxM, mnF, mxF, dz, peso, grp, fmt, orig, nota) in enumerate(IDX, 1):
    ws.cell(row=r, column=1, value=i).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=2, value=nome).font = Font(name=FONT, size=10, bold=True)
    for col, v in ((3, mnM), (4, mxM), (5, mnF), (6, mxF)):
        c = ws.cell(row=r, column=col, value=v)
        c.font = Font(name=FONT, size=10, color=BLUE)
        c.number_format = fmt
        c.fill = PatternFill("solid", fgColor=YEL)
        c.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=7, value=dz).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=8, value=peso).font = Font(name=FONT, size=10, color=BLUE)
    ws.cell(row=r, column=8).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=9, value=grp).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=10, value=orig).font = Font(name=FONT, size=9, italic=True, color="595959")
    diff = (mnM, mxM) != (mnF, mxF)
    for cc in range(1, 11):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        if cc in (2, 10):
            cell.alignment = Alignment(wrap_text=True, vertical="center")
    if diff:
        for cc in (3, 4, 5, 6):
            ws.cell(row=r, column=cc).font = Font(name=FONT, size=10, bold=True, color="C00000")
    ws.row_dimensions[r].height = 26
    r += 1
REF1 = r - 1
ws.freeze_panes = "C5"
nr = r + 1
ws.cell(row=nr, column=2, value="Le soglie in rosso sono quelle che cambiano fra uomo e donna. "
        "Sono %d indici su %d." % (sum(1 for x in IDX if (x[3], x[4]) != (x[5], x[6])), len(IDX))
        ).font = Font(name=FONT, size=10, bold=True, color="C00000")
ws.merge_cells(start_row=nr, start_column=2, end_row=nr, end_column=10)
ws.cell(row=nr + 1, column=2, value="A queste si aggiungono tre formule il cui calcolo stesso cambia per sesso: "
        "Visceral Adiposity Index (tutti i coefficienti), percentuale di grasso corporeo (termine sesso), "
        "peso ideale di Lorenz (denominatore) e frequenza cardiaca massima (Tanaka nell'uomo, Gulati nella donna)."
        ).font = Font(name=FONT, size=9, italic=True)
ws.merge_cells(start_row=nr + 1, start_column=2, end_row=nr + 2, end_column=10)
ws.cell(row=nr + 1, column=2).alignment = Alignment(wrap_text=True, vertical="top")

# =====================================================================
# 03 INDICI
# =====================================================================
ws = wb.create_sheet("03 Indici")
setw(ws, {"A": 5, "B": 40, "C": 12, "D": 11, "E": 10, "F": 10, "G": 14, "H": 8, "I": 8,
          "J": 10, "K": 16, "L": 58})
title(ws, "03 - Indici calcolati",
      "Tutto e calcolato dal foglio 01 con le soglie del foglio 02 secondo il sesso indicato. 'n.d.' significa che manca almeno un dato: quell'indice non entra nel punteggio.")
hdr(ws, 4, ["#", "Indice", "Valore", "Unita", "Min", "Max", "Stato", "Peso",
            "Punt.\n0-3", "Contrib.", "Gruppo", "Nota"])

R0 = 5
for i, (nome, unit, f, mnM, mxM, mnF, mxF, dz, peso, grp, fmt, orig, nota) in enumerate(IDX, 1):
    r = R0 + i - 1
    rr = REF0 + i - 1
    ws.cell(row=r, column=1, value=i).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=2, value=nome).font = Font(name=FONT, size=10, bold=True)
    c = ws.cell(row=r, column=3, value=f)
    c.font = Font(name=FONT, size=11, bold=True)
    c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=4, value=unit).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=5, value=f"=IF({SEX}=\"M\",'02 Riferimenti'!C{rr},'02 Riferimenti'!E{rr})").number_format = fmt
    ws.cell(row=r, column=6, value=f"=IF({SEX}=\"M\",'02 Riferimenti'!D{rr},'02 Riferimenti'!F{rr})").number_format = fmt
    ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=6).alignment = Alignment(horizontal="center")
    dzc = f"'02 Riferimenti'!$G${rr}"
    st = ws.cell(row=r, column=7, value=(
        f'=IF(NOT(ISNUMBER(C{r})),"n.d.",'
        f'IF({dzc}="Informativo",IF(OR(C{r}>F{r},C{r}<E{r}),"fuori fascia","in fascia"),'
        f'IF(AND(C{r}>F{r},OR({dzc}="Alto",{dzc}="Bilaterale")),"Fuori (alto)",'
        f'IF(AND(C{r}<E{r},OR({dzc}="Basso",{dzc}="Bilaterale")),"Fuori (basso)",'
        f'IF(OR(C{r}>F{r},C{r}<E{r}),"Fuori (non critico)",'
        f'IF(AND({dzc}<>"Informativo",F{r}<>E{r},OR(AND(OR({dzc}="Alto",{dzc}="Bilaterale"),(C{r}-E{r})/(F{r}-E{r})>0.8),'
        f'AND(OR({dzc}="Basso",{dzc}="Bilaterale"),(C{r}-E{r})/(F{r}-E{r})<0.2))),"Borderline","Normale"))))))'))
    st.font = Font(name=FONT, size=10, bold=True)
    st.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=8, value=f"='02 Riferimenti'!$H${rr}").alignment = Alignment(horizontal="center")
    pt = ws.cell(row=r, column=9, value=(
        f'=IF(OR(G{r}="n.d.",H{r}=0),"",'
        f'IF(AND(C{r}>F{r},OR({dzc}="Alto",{dzc}="Bilaterale")),IF((C{r}-F{r})/MAX(F{r}-E{r},0.0001)>0.1,3,2),'
        f'IF(AND(C{r}<E{r},OR({dzc}="Basso",{dzc}="Bilaterale")),IF((E{r}-C{r})/MAX(F{r}-E{r},0.0001)>0.1,3,2),'
        f'IF(G{r}="Borderline",1,0))))'))
    pt.font = Font(name=FONT, size=10, bold=True)
    pt.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=10, value=f'=IF(I{r}="","",H{r}*I{r})').alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=11, value=f"='02 Riferimenti'!$I${rr}").alignment = Alignment(horizontal="center")
    nc = ws.cell(row=r, column=12, value=nota)
    nc.font = Font(name=FONT, size=9)
    nc.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[r].height = rowh(nota, 74, 26, 66)
    for cc in range(1, 13):
        cell = ws.cell(row=r, column=cc)
        cell.border = BORDER
        if cell.font.name != FONT:
            cell.font = Font(name=FONT, size=10)
        if cc == 2:
            cell.alignment = Alignment(wrap_text=True, vertical="center")
LASTI = R0 + len(IDX) - 1
ws.freeze_panes = "C5"
ws.auto_filter.ref = f"A4:L{LASTI}"
rng = f"A{R0}:L{LASTI}"
ws.conditional_formatting.add(rng, FormulaRule(
    formula=[f'OR($G{R0}="Fuori (alto)",$G{R0}="Fuori (basso)")'], fill=BAD))
ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$G{R0}="Borderline"'], fill=WARN))
ws.conditional_formatting.add(rng, FormulaRule(formula=[f'$G{R0}="n.d."'],
                                               font=Font(name=FONT, size=10, color="A6A6A6")))

GRUPPI = sorted({x[9] for x in IDX})
ST = f"'03 Indici'!$G${R0}:$G${LASTI}"
GR = f"'03 Indici'!$K${R0}:$K${LASTI}"
CO = f"'03 Indici'!$J${R0}:$J${LASTI}"
PE = f"'03 Indici'!$H${R0}:$H${LASTI}"
PU = f"'03 Indici'!$I${R0}:$I${LASTI}"

# =====================================================================
# 04 PUNTEGGIO
# =====================================================================
ws = wb.create_sheet("04 Punteggio")
setw(ws, {"A": 4, "B": 34, "C": 13, "D": 13, "E": 13, "F": 13, "G": 62})
title(ws, "04 - Punteggio globale",
      "Il punteggio non somma i singoli indici: li raggruppa per misura di origine e prende la media di ciascun gruppo, cosi una sola misura alterata non conta piu volte.")

# riga del totale per gruppo, calcolata in anticipo perche il punteggio in cima la referenzia
G0_ROW = 17
G1_ROW = G0_ROW + len(GRUPPI) - 1
TOT_ROW = G1_ROW + 1

ws.cell(row=4, column=2, value="PUNTEGGIO GLOBALE").font = Font(name=FONT, size=12, bold=True, color="1F3864")
sc = ws.cell(row=4, column=3, value=f'=IF(OR(F{TOT_ROW}=0,F{TOT_ROW}=""),"n.d.",E{TOT_ROW}/F{TOT_ROW}*100)')
sc.font = Font(name=FONT, size=28, bold=True, color="C00000")
sc.number_format = "0.0"
sc.alignment = Alignment(horizontal="center", vertical="center")
ws.merge_cells(start_row=4, start_column=3, end_row=6, end_column=4)
fs = ws.cell(row=4, column=5, value=(
    '=IF(NOT(ISNUMBER(C4)),"dati insufficienti",'
    'IF(C11<0.4,"dati insufficienti",'
    'IF(C4<15,"BASSO",IF(C4<35,"LIEVE",IF(C4<55,"MODERATO",IF(C4<75,"ELEVATO","ALTO"))))))'))
fs.font = Font(name=FONT, size=18, bold=True, color="C00000")
fs.alignment = Alignment(horizontal="center", vertical="center")
ws.merge_cells(start_row=4, start_column=5, end_row=6, end_column=6)
ws.cell(row=4, column=7, value="Scala 0-100. Fasce: sotto 15 basso | 15-35 lieve | 35-55 moderato | 55-75 elevato | "
        "75 e oltre alto. Sotto il 40% di copertura la fascia non viene assegnata. Taratura "
        "verificata su profili di prova: un profilo sano da circa 2, uno marcatamente alterato circa 85. "
        "Resta una sintesi di orientamento, non uno score clinico validato.").font = Font(name=FONT, size=9, italic=True, color="595959")
ws.cell(row=4, column=7).alignment = Alignment(wrap_text=True, vertical="center")
ws.merge_cells(start_row=4, start_column=7, end_row=6, end_column=7)
for rr in range(4, 7):
    for cc in range(2, 8):
        ws.cell(row=rr, column=cc).border = BORDER

ws.cell(row=8, column=2, value="COMPLETEZZA DEI DATI").font = Font(name=FONT, size=11, bold=True, color="1F3864")
for k, (lab, f, fmt) in enumerate([
        ("Indici calcolabili con i dati inseriti", f'=COUNTIFS({ST},"<>n.d.")', "0"),
        ("Indici non calcolabili (dati mancanti)", f'=COUNTIF({ST},"n.d.")', "0"),
        ("Copertura", "=C9/(C9+C10)", "0%"),
        ("Indici fuori intervallo", f'=COUNTIF({ST},"Fuori (alto)")+COUNTIF({ST},"Fuori (basso)")', "0"),
        ("Indici in zona di guardia", f'=COUNTIF({ST},"Borderline")', "0")]):
    r = 9 + k
    ws.cell(row=r, column=2, value=lab).font = Font(name=FONT, size=10)
    c = ws.cell(row=r, column=3, value=f)
    c.font = Font(name=FONT, size=11, bold=True)
    c.number_format = fmt
    c.alignment = Alignment(horizontal="center")
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).border = BORDER
ws.cell(row=11, column=7, value="Sotto il 50% di copertura il punteggio va letto con molta cautela: "
        "misura solo la parte di quadro che i dati inseriti coprono.").font = Font(name=FONT, size=9, italic=True, color="595959")

ws.cell(row=15, column=2, value="COME SI FORMA IL PUNTEGGIO: contributo per gruppo di misura").font = Font(name=FONT, size=11, bold=True, color="1F3864")
hdr(ws, 16, ["", "Gruppo (misura di origine)", "Indici\ncalcolati", "Somma\ncontributi",
             "Contributo\nmedio", "Massimo\ndel gruppo", "Nota"], start=1, h=32)
r = G0_ROW
G0 = r
for g in GRUPPI:
    ws.cell(row=r, column=2, value=g).font = Font(name=FONT, size=10, bold=True)
    ws.cell(row=r, column=3, value=f'=COUNTIFS({GR},B{r},{PU},">=0")').alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=4, value=f'=SUMIFS({CO},{GR},B{r})').alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=5, value=f'=IF(C{r}=0,0,D{r}/C{r})').number_format = "0.00"
    ws.cell(row=r, column=6, value=f'=IF(C{r}=0,0,SUMIFS({PE},{GR},B{r},{PU},">=0")/C{r}*3)').number_format = "0.00"
    ws.cell(row=r, column=5).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=6).alignment = Alignment(horizontal="center")
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).border = BORDER
        if ws.cell(row=r, column=cc).font.name != FONT:
            ws.cell(row=r, column=cc).font = Font(name=FONT, size=10)
    r += 1
G1 = r - 1
ws.cell(row=G1 + 1, column=2, value="TOTALE").font = Font(name=FONT, size=10, bold=True)
ws.cell(row=G1 + 1, column=5, value=f"=SUM(E{G0}:E{G1})").font = Font(name=FONT, size=11, bold=True)
ws.cell(row=G1 + 1, column=6, value=f"=SUM(F{G0}:F{G1})").font = Font(name=FONT, size=11, bold=True)
for cc in range(2, 8):
    ws.cell(row=G1 + 1, column=cc).fill = SUB
    ws.cell(row=G1 + 1, column=cc).border = BORDER
ws.cell(row=G1 + 1, column=7, value="Il punteggio globale in cima e il rapporto fra queste due somme.").font = Font(name=FONT, size=9, italic=True)

assert G1 + 1 == TOT_ROW, f"riga totale attesa {TOT_ROW}, trovata {G1+1}"

ws.cell(row=G1 + 3, column=2, value="Perche la media e non la somma").font = Font(name=FONT, size=10, bold=True, color="C00000")
c = ws.cell(row=G1 + 4, column=2, value=(
    "Molti indici nascono dalla stessa misura: la sola circonferenza vita ne alimenta sei, il solo valore dei "
    "linfociti ne alimenta sette. Sommando i contributi, un paziente con un unico dato alterato riceverebbe un "
    "punteggio molto piu alto di quanto quel dato giustifichi. Raggruppando per misura di origine e prendendo la "
    "media di ogni gruppo, ogni misura pesa una volta sola, e il punteggio sale davvero solo quando sono alterate "
    "cose diverse fra loro."))
c.font = Font(name=FONT, size=9, italic=True)
c.alignment = Alignment(wrap_text=True, vertical="top")
ws.merge_cells(start_row=G1 + 4, start_column=2, end_row=G1 + 6, end_column=7)
for rr in range(G1 + 4, G1 + 7):
    for cc in range(2, 8):
        ws.cell(row=rr, column=cc).fill = WARN
        ws.cell(row=rr, column=cc).border = BORDER

# segnalazioni cliniche
sr = G1 + 8
ws.cell(row=sr, column=2, value="SEGNALAZIONI").font = Font(name=FONT, size=11, bold=True, color="1F3864")
SEG = [
    ("Fenotipo TOFI (normopeso con obesita centrale)",
     f'=IF(OR({BMI}="",{WC}="",{H}=""),"dati mancanti",IF(AND({BMI}<25,{WC}/{H}>0.5),"PRESENTE","assente"))',
     "BMI sotto 25 con vita oltre meta dell'altezza. La mortalita cardiovascolare supera quella dell'obeso con distribuzione periferica (PMID 26551006)."),
    ("Ferritina interpretabile",
     f'=IF(OR({PCR}="",{FER}=""),"dati mancanti",IF({PCR}>0.5,"NO - infiammazione in atto","si"))',
     "Con PCR alterata la ferritina misura anche l'infiammazione: va letta insieme alla saturazione della transferrina."),
    ("Pattern di ferro non disponibile",
     f'=IF(OR({FE}="",{TRF}="",{FER}=""),"dati mancanti",IF(AND({FE}/({TRF}*1.404)*100<20,{FER}>IF({SEX}="M",30,15)),"PRESENTE","assente"))',
     "Saturazione sotto il 20% con depositi non esauriti: il ferro c'e ma non arriva. In questo quadro la sola supplementazione marziale e spesso inefficace, la leva e la causa infiammatoria (PMID 30401705)."),
    ("Infiammazione cronica con indici classici negativi",
     f'=IF(OR({PCR}="",{NEU}="",{LYM}="",{MON}=""),"dati mancanti",IF(AND({PCR}<=0.5,{NEU}*{MON}/{LYM}>0.68),"PRESENTE","assente"))',
     "PCR normale ma SIRI oltre soglia: e la situazione che PCR e VES non vedono."),
    ("Sospetta apnea del sonno da approfondire",
     f'=IF({SNORE}="","non indagato",IF({SNORE}=1,"SI - proporre studio del sonno","no"))',
     "Domanda che non compare in nessun referto e che va posta a voce."),
    ("Sarcopenia da approfondire",
     f'=IF({CALF}="","dati mancanti",IF({CALF}<IF({SEX}="M",34,33),"SI - confermare con forza di presa","no"))',
     "Soglia 34 cm nell'uomo, 33 nella donna. Il polpaccio e un surrogato: la conferma richiede dinamometria o composizione corporea."),
]
r = sr + 1
for lab, f, nota in SEG:
    ws.cell(row=r, column=2, value=lab).font = Font(name=FONT, size=10)
    c = ws.cell(row=r, column=3, value=f)
    c.font = Font(name=FONT, size=10, bold=True, color="C00000")
    c.alignment = Alignment(horizontal="center")
    ws.merge_cells(start_row=r, start_column=3, end_row=r, end_column=4)
    n = ws.cell(row=r, column=7, value=nota)
    n.font = Font(name=FONT, size=9, italic=True, color="595959")
    n.alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = rowh(nota, 82, 26, 44)
    for cc in range(2, 8):
        ws.cell(row=r, column=cc).border = BORDER
    r += 1

# =====================================================================
# 05 FOLLOW-UP
# =====================================================================
ws = wb.create_sheet("05 Follow-up")
setw(ws, {"A": 4, "B": 34, "C": 14, "D": 14, "E": 14, "F": 14, "G": 14, "H": 13, "I": 46})
title(ws, "05 - Andamento nel tempo",
      "Riporta qui a mano i valori di ogni controllo. Diverse evidenze valgono sulla persistenza del dato, non su una misurazione isolata.")
hdr(ws, 4, ["", "Parametro", "Basale", "Controllo 1", "Controllo 2", "Controllo 3",
            "Controllo 4", "Variazione", "Nota"])
FU = [
    ("Data", "", "Compila con la data di ciascun controllo"),
    ("Punteggio globale", "0.0", "Copia il valore del foglio 04 a ogni controllo"),
    ("Peso (kg)", "0.0", ""),
    ("Circonferenza vita (cm)", "0.0", "La misura che si muove per prima con l'esercizio contro resistenza, spesso prima del peso"),
    ("Rapporto vita / altezza", "0.00", "Obiettivo: sotto 0,50"),
    ("Frequenza cardiaca a riposo (bpm)", "0", "Va confermata su piu misurazioni: l'HR 1,86 di Kailuan vale sulla persistenza sopra 80 bpm, non su un singolo prelievo"),
    ("Pressione sistolica (mmHg)", "0", ""),
    ("HDL (mg/dL)", "0", "Denominatore di quattro indici: si muove lentamente ma sposta molto"),
    ("Trigliceridi (mg/dL)", "0", ""),
    ("Saturazione transferrina (%)", "0.0", "Da rivalutare a tre mesi, meglio se insieme al contenuto emoglobinico reticolocitario"),
    ("Ferritina (ug/L)", "0", ""),
    ("PCR (mg/dL)", "0.00", ""),
    ("SIRI", "0.00", "Fra gli indici infiammatori e quello con lo scostamento piu leggibile"),
    ("Sigarette al giorno", "0", "La cessazione migliora la frequenza e la variabilita gia nelle prime settimane, ma il recupero pieno e piu lento"),
]
r = 5
for i, (lab, fmt, nota) in enumerate(FU, 1):
    ws.cell(row=r, column=1, value=i).alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=2, value=lab).font = Font(name=FONT, size=10, bold=True)
    for cc in range(3, 8):
        c = ws.cell(row=r, column=cc)
        c.fill = PatternFill("solid", fgColor=YEL)
        c.font = Font(name=FONT, size=10, color=BLUE)
        c.number_format = fmt if fmt else "General"
        c.alignment = Alignment(horizontal="center")
    if fmt:
        d = ws.cell(row=r, column=8, value=f'=IF(OR(C{r}="",COUNT(D{r}:G{r})=0),"",INDEX(D{r}:G{r},MATCH(9.99E+307,D{r}:G{r}))-C{r})')
        d.number_format = "+" + fmt + ";-" + fmt
        d.font = Font(name=FONT, size=10, bold=True)
        d.alignment = Alignment(horizontal="center")
    ws.cell(row=r, column=9, value=nota).font = Font(name=FONT, size=9, italic=True, color="595959")
    ws.cell(row=r, column=9).alignment = Alignment(wrap_text=True, vertical="center")
    ws.row_dimensions[r].height = rowh(nota, 60, 24, 44)
    for cc in range(1, 10):
        ws.cell(row=r, column=cc).border = BORDER
        if ws.cell(row=r, column=cc).font.name != FONT:
            ws.cell(row=r, column=cc).font = Font(name=FONT, size=10)
    r += 1
ws.cell(row=r + 1, column=2, value="La colonna 'Variazione' confronta l'ultimo controllo compilato con il basale.").font = Font(name=FONT, size=9, italic=True)

# =====================================================================
# 00 ISTRUZIONI
# =====================================================================
ws = wb.create_sheet("00 Istruzioni", 0)
setw(ws, {"A": 30, "B": 100})
title(ws, "Scheda indici cardiometabolici - versione per inserimento dati",
      "Uomini e donne. Formule e soglie si adattano automaticamente al sesso indicato nel foglio 01.")
IST = [
    ("", ""),
    ("A COSA SERVE", "Calcolare da misure di routine una batteria di indici che i referti ordinari non riportano, "
                     "e sintetizzarli in un punteggio unico di orientamento. Non richiede alcun esame aggiuntivo: "
                     "tutto si ricava da metro da sarta, bilancia, pressione, emocromo e un profilo biochimico di base."),
    ("", ""),
    ("COSA NON E", "Non e un dispositivo medico e non produce una diagnosi. Il punteggio e costruito su pesi scelti "
                   "dall'autore del foglio e non e uno score clinico validato su una popolazione. Serve a mettere in "
                   "fila dati che altrimenti restano sparsi, e a far emergere pattern che il singolo valore non mostra. "
                   "La decisione clinica resta interamente al medico che ha in cura la persona."),
    ("", ""),
    ("COME SI USA", "1) Foglio '01 Paziente': compila le celle gialle. Le grigie sono facoltative.\n"
                    "2) Indica il sesso: cambia le formule di quattro indici e le soglie di quindici.\n"
                    "3) Foglio '04 Punteggio': leggi il punteggio globale, la copertura dei dati e le segnalazioni.\n"
                    "4) Foglio '03 Indici': il dettaglio, se serve capire da dove viene il punteggio.\n"
                    "5) Foglio '05 Follow-up': riporta i controlli successivi."),
    ("", ""),
    ("DATI MANCANTI", "Un indice senza tutti i dati necessari mostra 'n.d.' e resta fuori dal punteggio: non viene "
                      "stimato ne sostituito. Il foglio 04 indica la percentuale di copertura raggiunta. "
                      "Sotto il 50% il punteggio va letto con molta cautela, perche misura solo la parte di quadro "
                      "che i dati coprono."),
    ("", ""),
    ("COSA NON SI PUO RICAVARE", "Alcuni dati non si ottengono per calcolo da altri, e il foglio non prova a inventarli: "
                                 "epcidina, contenuto emoglobinico reticolocitario, polisonnografia, test da sforzo, "
                                 "composizione corporea strumentale. Dove servirebbero, il foglio lo segnala come "
                                 "approfondimento da proporre, non come valore stimato."),
    ("", ""),
    ("STIME E MISURE", "Tre valori del foglio sono stime di popolazione, non misure: la percentuale di grasso corporeo "
                       "(derivata dal BMI), la frequenza cardiaca massima (derivata dall'eta) e il peso ideale. "
                       "Portano un errore che nelle formule successive non e piu visibile. Vanno trattati come "
                       "orientamento, non come dati del paziente."),
    ("", ""),
    ("IL PUNTEGGIO", "Non e la somma dei singoli indici. Molti indici nascono dalla stessa misura, quindi sommarli "
                     "conterebbe piu volte lo stesso dato: la circonferenza vita da sola alimenta sei indici, i "
                     "linfociti sette. Il punteggio raggruppa gli indici per misura di origine, fa la media di ogni "
                     "gruppo e somma i gruppi, cosi ogni misura pesa una volta sola. Il dettaglio e nel foglio 04."),
    ("", ""),
    ("SOGLIE", "Tutte le soglie stanno nel foglio '02 Riferimenti' e sono modificabili. Se il tuo laboratorio usa "
               "intervalli diversi, cambiali li: il resto del foglio si adegua. Le soglie evidenziate in rosso sono "
               "quelle che differiscono fra uomo e donna."),
    ("", ""),
    ("PRIVACY", "Il foglio non richiede dati identificativi. Se lo condividi, usa un codice al posto del nome."),
]
r = 4
for a, b in IST:
    ca, cb = ws.cell(row=r, column=1, value=a), ws.cell(row=r, column=2, value=b)
    ca.font = Font(name=FONT, size=10, bold=True, color="C00000" if a in ("COSA NON E", "COSA NON SI PUO RICAVARE") else "000000")
    cb.font = Font(name=FONT, size=10)
    cb.alignment = Alignment(wrap_text=True, vertical="top")
    ca.alignment = Alignment(vertical="top")
    if b:
        ws.row_dimensions[r].height = rowh(b, 100, 15, 76)
    r += 1

for sheet in wb.worksheets:
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value is not None and (cell.font is None or cell.font.name != FONT):
                f = cell.font
                cell.font = Font(name=FONT, size=f.size or 10, bold=f.bold, italic=f.italic,
                                 color=f.color, underline=f.underline)
    sheet.sheet_view.showGridLines = False

wb.save(OUT)
n_diff = sum(1 for x in IDX if (x[3], x[4]) != (x[5], x[6]))
print(f"scritto: {OUT}")
print(f"indici: {len(IDX)} | soglie diverse per sesso: {n_diff} | gruppi di ridondanza: {len(GRUPPI)}")
