from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.colors import HexColor, white, black
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, PageBreak, KeepTogether
)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_JUSTIFY
from reportlab.platypus import Flowable
import reportlab.lib.colors as colors

# ── Palette ──────────────────────────────────────────────────────────────────
NAVY      = HexColor("#0D1B2A")   # sfondo copertina / titoli principali
MIDNIGHT  = HexColor("#1B2A3B")   # sfondo box
ELECTRIC  = HexColor("#00B4D8")   # accent blu ghiaccio
GOLD      = HexColor("#F0A500")   # accent oro per checklist / numeri
CREAM     = HexColor("#F5F0E8")   # sfondo pagine interne
LIGHTGRAY = HexColor("#E8E8E8")
DARKTEXT  = HexColor("#1A1A2E")
MUTEDTEXT = HexColor("#555566")

W, H = A4  # 595 x 842 pt

# ── Documento ────────────────────────────────────────────────────────────────
doc = SimpleDocTemplate(
    "/home/user/prova/Protocollo_Spegnimento_Mentale.pdf",
    pagesize=A4,
    leftMargin=18*mm, rightMargin=18*mm,
    topMargin=12*mm, bottomMargin=12*mm,
)

styles = getSampleStyleSheet()

def style(name, **kw):
    return ParagraphStyle(name, **kw)

# ── Stili ─────────────────────────────────────────────────────────────────────
S = {
    "cover_title": style("ct", fontName="Helvetica-Bold", fontSize=28,
                         textColor=white, leading=34, alignment=TA_CENTER),
    "cover_sub":   style("cs", fontName="Helvetica", fontSize=13,
                         textColor=ELECTRIC, leading=18, alignment=TA_CENTER),
    "cover_tag":   style("ctag", fontName="Helvetica-Bold", fontSize=10,
                         textColor=GOLD, leading=14, alignment=TA_CENTER),
    "h1":          style("h1", fontName="Helvetica-Bold", fontSize=20,
                         textColor=NAVY, leading=26, spaceAfter=4),
    "h2":          style("h2", fontName="Helvetica-Bold", fontSize=14,
                         textColor=NAVY, leading=20, spaceBefore=8, spaceAfter=4),
    "h3":          style("h3", fontName="Helvetica-Bold", fontSize=11,
                         textColor=ELECTRIC, leading=16, spaceBefore=6, spaceAfter=2),
    "body":        style("body", fontName="Helvetica", fontSize=10.5,
                         textColor=DARKTEXT, leading=16, spaceAfter=4,
                         alignment=TA_JUSTIFY),
    "body_c":      style("body_c", fontName="Helvetica", fontSize=10.5,
                         textColor=DARKTEXT, leading=16, alignment=TA_CENTER),
    "quote":       style("quote", fontName="Helvetica-Oblique", fontSize=11,
                         textColor=white, leading=18, alignment=TA_CENTER),
    "box_title":   style("bt", fontName="Helvetica-Bold", fontSize=11,
                         textColor=GOLD, leading=16),
    "box_body":    style("bb", fontName="Helvetica", fontSize=10,
                         textColor=white, leading=15),
    "check":       style("ck", fontName="Helvetica", fontSize=10.5,
                         textColor=DARKTEXT, leading=18),
    "error_title": style("et", fontName="Helvetica-Bold", fontSize=11,
                         textColor=NAVY, leading=15),
    "error_body":  style("eb", fontName="Helvetica", fontSize=10,
                         textColor=MUTEDTEXT, leading=14),
    "phase_num":   style("pn", fontName="Helvetica-Bold", fontSize=36,
                         textColor=GOLD, leading=40, alignment=TA_CENTER),
    "phase_label": style("pl", fontName="Helvetica-Bold", fontSize=13,
                         textColor=ELECTRIC, leading=18, alignment=TA_CENTER),
    "small":       style("sm", fontName="Helvetica", fontSize=9,
                         textColor=MUTEDTEXT, leading=13),
    "footer":      style("ft", fontName="Helvetica", fontSize=8,
                         textColor=MUTEDTEXT, alignment=TA_CENTER),
    "formula_big": style("fb", fontName="Helvetica-Bold", fontSize=13,
                         textColor=NAVY, leading=20, alignment=TA_CENTER),
    "test_q":      style("tq", fontName="Helvetica", fontSize=10,
                         textColor=DARKTEXT, leading=15),
    "test_score":  style("ts", fontName="Helvetica-Bold", fontSize=10.5,
                         textColor=NAVY, leading=16),
}

# ── Helper Flowable: Colored Rectangle Background ────────────────────────────
class ColorRect(Flowable):
    def __init__(self, w, h, color, radius=4):
        super().__init__()
        self.w, self.h, self.color, self.radius = w, h, color, radius

    def draw(self):
        self.canv.setFillColor(self.color)
        self.canv.roundRect(0, 0, self.w, self.h, self.radius, fill=1, stroke=0)

    def wrap(self, *_):
        return self.w, self.h


def dark_box(content_paragraphs, bg=MIDNIGHT, pad=10):
    """Restituisce una Table con sfondo scuro."""
    inner = [p for p in content_paragraphs]
    t = Table([[inner]], colWidths=[W - 36*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), bg),
        ("ROUNDEDCORNERS",(0,0), (-1,-1), [6,6,6,6]),
        ("TOPPADDING",    (0,0), (-1,-1), pad),
        ("BOTTOMPADDING", (0,0), (-1,-1), pad),
        ("LEFTPADDING",   (0,0), (-1,-1), pad+2),
        ("RIGHTPADDING",  (0,0), (-1,-1), pad+2),
    ]))
    return t


def light_box(content_paragraphs, bg=LIGHTGRAY, pad=8):
    t = Table([[content_paragraphs]], colWidths=[W - 36*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",  (0,0), (-1,-1), bg),
        ("ROUNDEDCORNERS",(0,0),(-1,-1),[6,6,6,6]),
        ("TOPPADDING",  (0,0),(-1,-1), pad),
        ("BOTTOMPADDING",(0,0),(-1,-1), pad),
        ("LEFTPADDING", (0,0),(-1,-1), pad+2),
        ("RIGHTPADDING",(0,0),(-1,-1), pad+2),
    ]))
    return t


def checklist_table(items):
    """Genera una tabella checklist con simbolo □ dorato."""
    rows = []
    for item in items:
        rows.append([
            Paragraph("<font color='#F0A500'><b>□</b></font>", S["check"]),
            Paragraph(item, S["check"]),
        ])
    t = Table(rows, colWidths=[12*mm, W - 36*mm - 12*mm])
    t.setStyle(TableStyle([
        ("VALIGN",      (0,0),(-1,-1), "TOP"),
        ("LEFTPADDING", (0,0),(0,-1),  0),
        ("RIGHTPADDING",(0,0),(0,-1),  4),
        ("LEFTPADDING", (1,0),(1,-1),  0),
        ("TOPPADDING",  (0,0),(-1,-1), 2),
        ("BOTTOMPADDING",(0,0),(-1,-1),2),
    ]))
    return t


def numbered_list(items):
    rows = []
    for i, item in enumerate(items, 1):
        rows.append([
            Paragraph(f"<font color='#F0A500'><b>{i}</b></font>", S["body"]),
            Paragraph(item, S["body"]),
        ])
    t = Table(rows, colWidths=[10*mm, W - 36*mm - 10*mm])
    t.setStyle(TableStyle([
        ("VALIGN",       (0,0),(-1,-1),"TOP"),
        ("LEFTPADDING",  (0,0),(0,-1), 0),
        ("RIGHTPADDING", (0,0),(0,-1), 6),
        ("LEFTPADDING",  (1,0),(1,-1), 0),
        ("TOPPADDING",   (0,0),(-1,-1),2),
        ("BOTTOMPADDING",(0,0),(-1,-1),2),
    ]))
    return t


def divider(color=ELECTRIC, thick=1.5):
    return HRFlowable(width="100%", thickness=thick, color=color,
                      spaceAfter=6, spaceBefore=6)


def sp(n=6):
    return Spacer(1, n)


# ═════════════════════════════════════════════════════════════════════════════
# COPERTINA
# ═════════════════════════════════════════════════════════════════════════════
def build_cover():
    elems = []

    # Sfondo navy pieno (simulato con una grande tabella)
    cover_bg = Table(
        [[" "]],
        colWidths=[W - 36*mm],
        rowHeights=[H - 80*mm],
    )
    cover_bg.setStyle(TableStyle([
        ("BACKGROUND", (0,0),(-1,-1), NAVY),
        ("ROUNDEDCORNERS",(0,0),(-1,-1),[10,10,10,10]),
    ]))
    elems.append(cover_bg)

    # Overlay contenuto sulla copertina (ricostruito come tabella nested)
    cover_content = [
        sp(20),
        Paragraph("MENTAL PRO", S["cover_tag"]),
        sp(8),
        divider(GOLD, 2),
        sp(8),
        Paragraph("Protocollo di<br/>Spegnimento Mentale™", S["cover_title"]),
        sp(10),
        Paragraph("Il protocollo scientifico di 20 minuti per disattivare<br/>"
                  "i loop mentali e preparare il cervello al sonno profondo",
                  S["cover_sub"]),
        sp(20),
        divider(ELECTRIC, 1),
        sp(10),
        Paragraph("20 MIN &nbsp;·&nbsp; 3 FASI &nbsp;·&nbsp; RISULTATI DAL PRIMO UTILIZZO",
                  S["cover_tag"]),
        sp(8),
    ]

    cover_main = Table(
        [[cover_content]],
        colWidths=[W - 36*mm],
        rowHeights=[H - 80*mm],
    )
    cover_main.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), NAVY),
        ("ROUNDEDCORNERS",(0,0),(-1,-1),[10,10,10,10]),
        ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",  (0,0),(-1,-1), 16),
        ("RIGHTPADDING", (0,0),(-1,-1), 16),
        ("TOPPADDING",   (0,0),(-1,-1), 0),
        ("BOTTOMPADDING",(0,0),(-1,-1), 0),
    ]))

    # Rimpiazza il precedente cover_bg con il vero contenuto
    elems = [cover_main]
    elems.append(sp(8))
    elems.append(Paragraph("© Mental Pro – Uso personale. Riservato ai clienti.",
                            S["footer"]))
    elems.append(PageBreak())
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# INTRODUZIONE
# ═════════════════════════════════════════════════════════════════════════════
def build_intro():
    elems = []
    elems.append(Paragraph("Introduzione", S["h1"]))
    elems.append(divider())
    elems.append(sp(4))

    elems.append(dark_box([
        Paragraph("Ti è mai capitato di essere fisicamente stanco<br/>"
                  "ma mentalmente ancora in riunione?", S["quote"]),
    ]))
    elems.append(sp(10))

    elems.append(Paragraph(
        "Se la risposta è sì, non hai un problema di sonno. Hai un problema di <b>modalità cognitiva</b>. "
        "Il tuo cervello è ancora acceso, ancora in analisi, ancora a cercare soluzioni a problemi che "
        "domani mattina esistono ancora.",
        S["body"]
    ))
    elems.append(sp(6))
    elems.append(Paragraph(
        "La maggior parte delle persone crede che dormire male dipenda dallo stress, dall'ansia o da "
        "qualcosa di incontrollabile. In realtà, nella maggior parte dei casi dipende da un fatto "
        "molto più semplice: <b>non è mai stato inviato al cervello il segnale che la giornata è finita</b>.",
        S["body"]
    ))
    elems.append(sp(10))

    elems.append(light_box([
        Paragraph("Il principio fondamentale", S["box_title"]),
        sp(4),
        Paragraph(
            "Non devi spegnere il corpo. Devi spegnere il <b>circuito mentale</b> "
            "che continua a cercare problemi da risolvere.",
            style("lbb", fontName="Helvetica", fontSize=10.5,
                  textColor=DARKTEXT, leading=16)
        ),
    ], bg=HexColor("#EAF4FB")))

    elems.append(sp(10))
    elems.append(Paragraph(
        "Questo protocollo è stato progettato per lavorare in sinergia con Mental Pro. "
        "Il supplemento agisce a livello biochimico; il rituale agisce a livello cognitivo e sensoriale. "
        "Insieme, creano le condizioni ottimali per interrompere il rimuginamento serale e "
        "permettere al cervello di entrare in modalità recupero.",
        S["body"]
    ))
    elems.append(sp(6))
    elems.append(Paragraph(
        "<b>Tempo totale: 20 minuti.</b> Strutturati in 3 fasi. "
        "Non richiedono attrezzature, app, o discipline speciali.",
        S["body"]
    ))
    elems.append(PageBreak())
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# FASE HEADER (funzione riutilizzabile)
# ═════════════════════════════════════════════════════════════════════════════
def phase_header(number, title, time_label, objective):
    rows = [[
        Paragraph(str(number), S["phase_num"]),
        [
            Paragraph(f"FASE {number}", S["phase_label"]),
            Paragraph(f"<b>{title}</b>", style("phT", fontName="Helvetica-Bold",
                      fontSize=16, textColor=white, leading=22)),
            Paragraph(f"⏱ {time_label}", style("phTm", fontName="Helvetica",
                      fontSize=10, textColor=ELECTRIC, leading=14)),
        ]
    ]]
    t = Table(rows, colWidths=[20*mm, W - 36*mm - 20*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), NAVY),
        ("ROUNDEDCORNERS",(0,0),(-1,-1),[8,8,8,8]),
        ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
        ("LEFTPADDING",  (0,0),(0,-1), 10),
        ("RIGHTPADDING", (0,0),(0,-1), 6),
        ("LEFTPADDING",  (1,0),(1,-1), 8),
        ("TOPPADDING",   (0,0),(-1,-1),12),
        ("BOTTOMPADDING",(0,0),(-1,-1),12),
    ]))
    elems = [t, sp(8)]
    elems.append(dark_box([
        Paragraph("Obiettivo", S["box_title"]),
        sp(3),
        Paragraph(objective, S["box_body"]),
    ]))
    elems.append(sp(10))
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# FASE 1
# ═════════════════════════════════════════════════════════════════════════════
def build_fase1():
    elems = []
    elems += phase_header(1, "Chiusura Cognitiva", "5 minuti",
                          'Comunicare al cervello: "La giornata è finita."')

    elems.append(Paragraph("Checklist", S["h3"]))
    elems.append(checklist_table([
        "Scrivi su carta le <b>3 priorità di domani</b>",
        "Scrivi l'<b>ultima attività incompleta</b>",
        "Definisci il <b>primo passo concreto</b> per completarla",
        "Chiudi agenda e computer",
    ]))
    elems.append(sp(12))

    elems.append(dark_box([
        Paragraph("Perché funziona — Box Scientifico", S["box_title"]),
        sp(4),
        Paragraph(
            "Il cervello mantiene attivi i compiti incompleti attraverso un meccanismo "
            "chiamato <b>Effetto Zeigarnik</b>: finché un'attività non è risolta, "
            "il sistema cognitivo la tiene in memoria di lavoro, occupando risorse mentali "
            "che impediscono il rilassamento.",
            S["box_body"]
        ),
        sp(6),
        Paragraph(
            "Quando trasferisci il compito su carta, il cervello <i>percepisce</i> che esiste "
            "un sistema esterno che lo gestirà. Questo riduce la necessità di mantenerlo attivo, "
            "liberando le risorse necessarie per il passaggio allo stato di riposo.",
            S["box_body"]
        ),
    ]))
    elems.append(sp(6))

    elems.append(Paragraph(
        "<i>Nota: non devi risolvere nulla in questo momento. "
        "L'obiettivo è solo trasferire il compito fuori dalla mente.</i>",
        S["small"]
    ))
    elems.append(PageBreak())
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# FASE 2
# ═════════════════════════════════════════════════════════════════════════════
def build_fase2():
    elems = []
    elems += phase_header(2, "Disconnessione Neuro-Sensoriale", "10 minuti",
                          "Ridurre gli stimoli che mantengono acceso il sistema nervoso.")

    elems.append(Paragraph("Checklist", S["h3"]))
    elems.append(checklist_table([
        "Telefono in <b>modalità aereo</b>",
        "Nessuna email — né in lettura né in risposta",
        "Nessuna notizia (news, social, feed)",
        "Nessun contenuto lavorativo o di produttività",
        "Luci abbassate (ideale: luce calda sotto i 2700K)",
        "Temperatura ambiente leggermente più fresca",
    ]))
    elems.append(sp(12))

    elems.append(dark_box([
        Paragraph("Regola Mental Pro", S["box_title"]),
        sp(4),
        Paragraph(
            '"Se un contenuto può generare una decisione,<br/>non guardarlo.'
            ' Nemmeno per un secondo."',
            style("rmp", fontName="Helvetica-Oblique", fontSize=11,
                  textColor=GOLD, leading=18, alignment=TA_CENTER)
        ),
    ]))
    elems.append(sp(10))

    elems.append(Paragraph(
        "Il sistema nervoso autonomo non distingue tra una minaccia reale e uno stimolo digitale "
        "percepito come urgente. Un'email non letta, una notizia preoccupante, una notifica "
        "irrisolta: ciascuno di questi stimoli mantiene attivo il cortisolo e impedisce "
        "l'abbassamento della frequenza cerebrale verso le onde alpha e theta, necessarie "
        "per l'addormentamento.",
        S["body"]
    ))
    elems.append(PageBreak())
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# FASE 3
# ═════════════════════════════════════════════════════════════════════════════
def build_fase3():
    elems = []
    elems += phase_header(3, "Attivazione del Protocollo Calmante", "5 minuti",
                          "Favorire il passaggio dalla modalità 'performance' alla modalità 'recupero'.")

    elems.append(Paragraph("Checklist", S["h3"]))
    elems.append(checklist_table([
        "Assunzione di <b>Mental Pro</b>",
        "Respirazione lenta: 4 secondi inspira, 6 secondi espira — per 2 minuti",
        "Lettura leggera (narrativa, non business)",
        "Nessuna pianificazione o problem solving",
        "Nessuna discussione impegnativa",
    ]))
    elems.append(sp(12))

    elems.append(light_box([
        Paragraph("Nota importante", style("ni", fontName="Helvetica-Bold",
                  fontSize=10.5, textColor=NAVY, leading=15)),
        sp(3),
        Paragraph(
            "Non stai cercando di addormentarti. "
            "Stai creando le condizioni perché il cervello smetta di restare in allerta. "
            "Il sonno arriva da solo quando il sistema nervoso riceve segnali coerenti di sicurezza.",
            style("nib", fontName="Helvetica", fontSize=10.5,
                  textColor=DARKTEXT, leading=16)
        ),
    ], bg=HexColor("#FFF8E7")))

    elems.append(PageBreak())
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# 5 ERRORI
# ═════════════════════════════════════════════════════════════════════════════
def build_errori():
    elems = []
    elems.append(Paragraph("I 5 Errori che Riaccendono il Cervello", S["h1"]))
    elems.append(divider())
    elems.append(sp(6))
    elems.append(Paragraph(
        "Puoi seguire il protocollo alla perfezione, ma questi comportamenti "
        "rischiano di azzerare i progressi. Evitarli è tanto importante quanto seguire le 3 fasi.",
        S["body"]
    ))
    elems.append(sp(10))

    errori = [
        ("Controllare le email dopo cena",
         "Anche una sola email non risposta crea un loop cognitivo aperto. "
         "Il cervello inizierà a elaborare la risposta mentre cerchi di dormire."),
        ("Guardare video di business o produttività",
         "Il cervello non distingue tra consumo passivo e attivazione cognitiva. "
         "I contenuti informativi densi mantengono il sistema prefrontale attivo."),
        ("Sforzarsi di dormire",
         "Cercare attivamente di addormentarsi crea una risposta da 'performance anxiety' "
         "che eleva il cortisolo. Il sonno richiede assenza di sforzo, non presenza di forza."),
        ("Risollevare problemi di lavoro a letto",
         "Il letto deve essere associato esclusivamente al riposo. "
         "Ogni discussione impegnativa rinforza l'associazione letto = stress."),
        ("Usare il telefono come ultima attività",
         "La luce blu sopprime la melatonina. Ma il problema principale non è ottico: "
         "è cognitivo. Il feed è progettato per generare stimoli infiniti e incompletezza."),
    ]

    for i, (titolo, desc) in enumerate(errori, 1):
        row_content = [
            Paragraph(f"<font color='#F0A500'><b>{i}</b></font>", S["phase_num"]),
            [
                Paragraph(titolo, S["error_title"]),
                Paragraph(desc, S["error_body"]),
            ]
        ]
        t = Table([row_content], colWidths=[16*mm, W - 36*mm - 16*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), LIGHTGRAY),
            ("ROUNDEDCORNERS",(0,0),(-1,-1),[6,6,6,6]),
            ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",  (0,0),(-1,-1), 8),
            ("RIGHTPADDING", (0,0),(-1,-1), 8),
            ("TOPPADDING",   (0,0),(-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1), 8),
        ]))
        elems.append(t)
        elems.append(sp(6))

    elems.append(PageBreak())
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# PROTOCOLLO EMERGENZA
# ═════════════════════════════════════════════════════════════════════════════
def build_emergenza():
    elems = []
    elems.append(Paragraph("Protocollo Rapido d'Emergenza", S["h1"]))
    elems.append(divider())
    elems.append(sp(4))

    elems.append(dark_box([
        Paragraph("Quando sei già a letto e la mente corre", S["box_title"]),
    ]))
    elems.append(sp(10))

    elems.append(Paragraph(
        "Questo protocollo si attiva solo quando sei già sdraiato e il cervello "
        "si è riacceso. Non è una sostituzione delle 3 fasi: è il piano B per "
        "le notti più difficili.",
        S["body"]
    ))
    elems.append(sp(10))

    elems.append(numbered_list([
        "<b>Non guardare l'orologio</b> — misurare l'insonnia la peggiora",
        "<b>Non controllare il telefono</b> — nemmeno per un secondo",
        "<b>Non cercare di forzare il sonno</b> — rilassa l'intenzione",
        "<b>Focalizzati solo sul respiro</b> — conta le espirazioni fino a 10, poi ricomincia",
        "<b>Lascia passare il pensiero senza completarlo</b> — osservalo come una nuvola",
    ]))
    elems.append(sp(12))

    elems.append(light_box([
        Paragraph(
            "Il cervello in modalità overthinking cerca una soluzione. "
            "La risposta corretta è non cercarla. "
            "Ogni tentativo di risolvere a letto prolunga l'attivazione.",
            style("emq", fontName="Helvetica-Oblique", fontSize=11,
                  textColor=DARKTEXT, leading=17, alignment=TA_CENTER)
        ),
    ], bg=HexColor("#EAF4FB")))

    elems.append(PageBreak())
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# TEST DI AUTOVALUTAZIONE
# ═════════════════════════════════════════════════════════════════════════════
def build_test():
    elems = []
    elems.append(Paragraph("Test di Autovalutazione", S["h1"]))
    elems.append(divider())
    elems.append(sp(4))
    elems.append(Paragraph(
        "Quanto è acceso il tuo cervello prima di dormire?",
        S["h2"]
    ))
    elems.append(Paragraph(
        "Per ogni domanda assegna un punteggio da <b>1</b> (mai/raramente) a <b>4</b> (sempre/quasi sempre).",
        S["body"]
    ))
    elems.append(sp(8))

    domande = [
        "La sera continui a pensare a cose che devi fare domani.",
        "Controlli email o messaggi di lavoro dopo le 20:00.",
        "Hai difficoltà a 'staccare' mentalmente anche quando sei fisicamente a casa.",
        "I pensieri di lavoro ti seguono fino al momento in cui chiudi gli occhi.",
        "Usi il telefono come ultima attività prima di dormire.",
        "Ti svegli durante la notte con pensieri ricorrenti.",
        "Ti ci vogliono più di 20 minuti per addormentarti.",
        "Senti che la tua mente 'analizza' anche quando non vuoi.",
        "Hai difficoltà a rilassarti senza un'attività da fare.",
        "Il giorno dopo senti stanchezza mentale anche dopo 7-8 ore di sonno.",
    ]

    for i, d in enumerate(domande, 1):
        rows = [[
            Paragraph(f"<b>{i}</b>", S["test_score"]),
            Paragraph(d, S["test_q"]),
            Paragraph("[ 1 ]  [ 2 ]  [ 3 ]  [ 4 ]",
                      style("sc", fontName="Helvetica", fontSize=9,
                            textColor=MUTEDTEXT, leading=13)),
        ]]
        t = Table(rows, colWidths=[8*mm, W - 36*mm - 8*mm - 35*mm, 35*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), CREAM if i % 2 == 0 else white),
            ("VALIGN",       (0,0),(-1,-1), "MIDDLE"),
            ("LEFTPADDING",  (0,0),(-1,-1), 6),
            ("RIGHTPADDING", (0,0),(-1,-1), 6),
            ("TOPPADDING",   (0,0),(-1,-1), 5),
            ("BOTTOMPADDING",(0,0),(-1,-1), 5),
            ("LINEBELOW",    (0,0),(-1,-1), 0.5, LIGHTGRAY),
        ]))
        elems.append(t)

    elems.append(sp(14))
    elems.append(Paragraph("Leggi il tuo punteggio", S["h2"]))
    elems.append(sp(4))

    punteggi = [
        ("10–16", "Cervello già rilassato",
         "Il tuo rituale serale funziona bene. Mental Pro ti aiuterà a consolidare ulteriormente la qualità del sonno."),
        ("17–24", "Attivazione moderata",
         "Alcune abitudini serali mantengono il cervello parzialmente attivo. Le 3 fasi del protocollo faranno una differenza significativa."),
        ("25–32", "Modalità Overthinking",
         "Il tuo sistema cognitivo rimane in allerta quasi ogni sera. Il Protocollo di Spegnimento è esattamente quello di cui hai bisogno."),
        ("33–40", "Cervello in Overdrive",
         "Il ciclo di analisi serale è molto intenso. Inizia dal Protocollo per 7 giorni consecutivi e osserva i cambiamenti."),
    ]

    for punteggio, label, desc in punteggi:
        t = Table([[
            Paragraph(f"<b>{punteggio}</b>", style("pt", fontName="Helvetica-Bold",
                      fontSize=13, textColor=GOLD, leading=18, alignment=TA_CENTER)),
            [
                Paragraph(f"<b>{label}</b>", S["test_score"]),
                Paragraph(desc, S["test_q"]),
            ]
        ]], colWidths=[20*mm, W - 36*mm - 20*mm])
        t.setStyle(TableStyle([
            ("BACKGROUND",   (0,0),(-1,-1), MIDNIGHT),
            ("ROUNDEDCORNERS",(0,0),(-1,-1),[6,6,6,6]),
            ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",  (0,0),(-1,-1), 8),
            ("RIGHTPADDING", (0,0),(-1,-1), 8),
            ("TOPPADDING",   (0,0),(-1,-1), 8),
            ("BOTTOMPADDING",(0,0),(-1,-1), 8),
        ]))
        elems.append(t)
        elems.append(sp(6))

    elems.append(PageBreak())
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# PAGINA FINALE
# ═════════════════════════════════════════════════════════════════════════════
def build_finale():
    elems = []
    elems.append(Paragraph("La Formula Completa", S["h1"]))
    elems.append(divider())
    elems.append(sp(10))

    # Formula box
    formula = Table([[
        Paragraph("Mental Pro", style("fp1", fontName="Helvetica-Bold", fontSize=16,
                  textColor=GOLD, leading=22, alignment=TA_CENTER)),
        Paragraph("+", style("fpl", fontName="Helvetica-Bold", fontSize=22,
                  textColor=white, leading=28, alignment=TA_CENTER)),
        Paragraph("Rituale di<br/>Spegnimento™", style("fp2", fontName="Helvetica-Bold",
                  fontSize=14, textColor=ELECTRIC, leading=20, alignment=TA_CENTER)),
        Paragraph("=", style("fpl2", fontName="Helvetica-Bold", fontSize=22,
                  textColor=white, leading=28, alignment=TA_CENTER)),
        Paragraph("Condizioni<br/>Ottimali", style("fp3", fontName="Helvetica-Bold",
                  fontSize=14, textColor=white, leading=20, alignment=TA_CENTER)),
    ]], colWidths=[(W-36*mm)*0.28, (W-36*mm)*0.08, (W-36*mm)*0.28,
                  (W-36*mm)*0.08, (W-36*mm)*0.28])
    formula.setStyle(TableStyle([
        ("BACKGROUND",   (0,0),(-1,-1), NAVY),
        ("ROUNDEDCORNERS",(0,0),(-1,-1),[8,8,8,8]),
        ("VALIGN",       (0,0),(-1,-1),"MIDDLE"),
        ("TOPPADDING",   (0,0),(-1,-1),16),
        ("BOTTOMPADDING",(0,0),(-1,-1),16),
        ("LEFTPADDING",  (0,0),(-1,-1), 6),
        ("RIGHTPADDING", (0,0),(-1,-1), 6),
    ]))
    elems.append(formula)
    elems.append(sp(16))

    elems.append(Paragraph(
        "per interrompere il rimuginamento serale",
        S["body_c"]
    ))
    elems.append(sp(20))
    elems.append(divider(GOLD))
    elems.append(sp(14))

    elems.append(dark_box([
        sp(6),
        Paragraph(
            "Le persone ad alte prestazioni non dormono bene<br/>perché sono meno stressate.",
            S["quote"]
        ),
        sp(8),
        Paragraph(
            "Dormono bene perché sanno<br/><b>quando smettere di lavorare</b>.",
            style("qb", fontName="Helvetica-Bold", fontSize=13,
                  textColor=GOLD, leading=20, alignment=TA_CENTER)
        ),
        sp(6),
    ], bg=NAVY))

    elems.append(sp(20))
    elems.append(Paragraph(
        "Usa questo protocollo ogni sera per almeno 7 giorni. "
        "La costanza è il moltiplicatore.",
        S["body_c"]
    ))
    elems.append(sp(20))
    elems.append(divider(LIGHTGRAY, 0.5))
    elems.append(sp(6))
    elems.append(Paragraph(
        "© Mental Pro &nbsp;·&nbsp; Guida riservata ai clienti &nbsp;·&nbsp; Non distribuire",
        S["footer"]
    ))
    return elems


# ═════════════════════════════════════════════════════════════════════════════
# ASSEMBLY
# ═════════════════════════════════════════════════════════════════════════════
story = []
story += build_cover()
story += build_intro()
story += build_fase1()
story += build_fase2()
story += build_fase3()
story += build_errori()
story += build_emergenza()
story += build_test()
story += build_finale()

doc.build(story)
print("PDF generato: Protocollo_Spegnimento_Mentale.pdf")
