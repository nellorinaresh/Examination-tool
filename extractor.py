#!/usr/bin/env python3
"""
Charge Sheet Analyser — AI Powered (Multi-Turn)
IPC/CrPC & BNS/BNSS | All Courts | Magistrate & Sessions
Powered by Anthropic Claude API — Multi-Turn for Accuracy
"""

import threading, os, sys, zipfile, re, json, io, base64, copy
from datetime import date
from pathlib import Path

# ── Library checks ─────────────────────────────────────────────
MISSING = []
try:
    import fitz; HAS_PDF = True
except ImportError:
    HAS_PDF = False; MISSING.append("PyMuPDF")

try:
    from docx import Document as DocxDoc
    from docx.shared import Pt, Cm
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False; MISSING.append("python-docx")

try:
    import anthropic; HAS_AI = True
except ImportError:
    HAS_AI = False; MISSING.append("anthropic")

try:
    import mammoth; HAS_MAMMOTH = True
except ImportError:
    HAS_MAMMOTH = False

# ═══════════════════════════════════════════════════════════════
# SECTION DATABASE
# ═══════════════════════════════════════════════════════════════
DB = {
    "IPC":{
        "279":{"d":"rash driving or riding on a public way","m":6},
        "302":{"d":"murder","m":9999},
        "304":{"d":"culpable homicide not amounting to murder","m":120},
        "304A":{"d":"causing death by negligence","m":24},
        "304(A)":{"d":"causing death by negligence","m":24},
        "304B":{"d":"dowry death","m":9999},
        "304(B)":{"d":"dowry death","m":9999},
        "306":{"d":"abetment of suicide","m":120},
        "307":{"d":"attempt to murder","m":120},
        "308":{"d":"attempt to commit culpable homicide","m":36},
        "323":{"d":"voluntarily causing hurt","m":12},
        "324":{"d":"voluntarily causing hurt by dangerous weapons","m":36},
        "325":{"d":"voluntarily causing grievous hurt","m":84},
        "326":{"d":"voluntarily causing grievous hurt by dangerous weapons","m":9999},
        "326A":{"d":"voluntarily causing grievous hurt by use of acid","m":9999},
        "341":{"d":"wrongful restraint","m":1},
        "342":{"d":"wrongful confinement","m":12},
        "354":{"d":"assault or criminal force to woman to outrage modesty","m":24},
        "354A":{"d":"sexual harassment","m":36},
        "354B":{"d":"assault with intent to disrobe","m":84},
        "354C":{"d":"voyeurism","m":84},
        "354D":{"d":"stalking","m":60},
        "363":{"d":"kidnapping","m":84},
        "364":{"d":"kidnapping for murder","m":9999},
        "364A":{"d":"kidnapping for ransom","m":9999},
        "366":{"d":"kidnapping woman to compel marriage","m":120},
        "376":{"d":"rape","m":9999},
        "376(1)":{"d":"rape","m":9999},
        "376(2)":{"d":"aggravated rape","m":9999},
        "376A":{"d":"rape causing death or vegetative state","m":9999},
        "376D":{"d":"gang rape","m":9999},
        "379":{"d":"theft","m":36},
        "380":{"d":"theft in dwelling house","m":84},
        "392":{"d":"robbery","m":120},
        "395":{"d":"dacoity","m":9999},
        "396":{"d":"dacoity with murder","m":9999},
        "406":{"d":"criminal breach of trust","m":36},
        "409":{"d":"criminal breach of trust by public servant","m":9999},
        "420":{"d":"cheating and dishonestly inducing delivery of property","m":84},
        "427":{"d":"mischief causing damage","m":24},
        "435":{"d":"mischief by fire to damage property","m":84},
        "436":{"d":"mischief by fire to destroy house","m":9999},
        "447":{"d":"criminal trespass","m":3},
        "448":{"d":"house-trespass","m":12},
        "452":{"d":"house-trespass after preparation for hurt","m":84},
        "457":{"d":"lurking house-trespass by night","m":84},
        "465":{"d":"forgery","m":24},
        "467":{"d":"forgery of valuable security or will","m":9999},
        "468":{"d":"forgery for purpose of cheating","m":84},
        "471":{"d":"using as genuine a forged document","m":24},
        "489A":{"d":"counterfeiting currency notes","m":9999},
        "489B":{"d":"using as genuine forged currency notes","m":9999},
        "498A":{"d":"husband or relative subjecting woman to cruelty","m":36},
        "498(A)":{"d":"husband or relative subjecting woman to cruelty","m":36},
        "504":{"d":"intentional insult to provoke breach of peace","m":24},
        "506":{"d":"criminal intimidation","m":24},
        "506(1)":{"d":"criminal intimidation","m":24},
        "506(2)":{"d":"criminal intimidation with threat of death or grievous hurt","m":84},
        "509":{"d":"word or gesture to insult modesty of woman","m":36},
        "34":{"d":"acts done in furtherance of common intention","m":0},
        "149":{"d":"every member of unlawful assembly guilty of offence","m":0},
        "120B":{"d":"criminal conspiracy","m":9999},
        "147":{"d":"rioting","m":24},
        "148":{"d":"rioting armed with deadly weapon","m":36},
        "186":{"d":"obstructing public servant","m":3},
        "188":{"d":"disobedience to order by public servant","m":6},
        "193":{"d":"giving false evidence","m":84},
        "201":{"d":"causing disappearance of evidence","m":84},
    },
    "BNS":{
        "103":{"d":"murder","m":9999},
        "105":{"d":"culpable homicide not amounting to murder","m":120},
        "106":{"d":"causing death by negligence","m":60},
        "106(1)":{"d":"causing death by negligence","m":60},
        "106(2)":{"d":"causing death by negligence — hit and run","m":120},
        "108":{"d":"abetment of suicide","m":120},
        "109":{"d":"attempt to murder","m":120},
        "115":{"d":"voluntarily causing hurt","m":12},
        "115(1)":{"d":"voluntarily causing hurt","m":12},
        "115(2)":{"d":"voluntarily causing grievous hurt","m":84},
        "116":{"d":"voluntarily causing hurt by dangerous weapons","m":36},
        "116(1)":{"d":"voluntarily causing hurt by dangerous weapons","m":36},
        "116(2)":{"d":"voluntarily causing grievous hurt by dangerous weapons","m":9999},
        "117":{"d":"voluntarily causing grievous hurt","m":84},
        "118":{"d":"voluntarily causing grievous hurt by dangerous weapons","m":9999},
        "118(1)":{"d":"voluntarily causing grievous hurt by dangerous weapons or means","m":9999},
        "118(2)":{"d":"voluntarily causing grievous hurt by dangerous weapons causing death","m":9999},
        "121":{"d":"assault or criminal force to woman to outrage modesty","m":24},
        "122":{"d":"sexual harassment","m":36},
        "123":{"d":"assault with intent to disrobe woman","m":84},
        "124":{"d":"voyeurism","m":84},
        "125":{"d":"stalking","m":60},
        "126":{"d":"wrongful confinement","m":12},
        "128":{"d":"kidnapping","m":84},
        "130":{"d":"kidnapping for murder","m":9999},
        "131":{"d":"kidnapping for ransom","m":9999},
        "64":{"d":"rape","m":9999},
        "64(1)":{"d":"rape","m":9999},
        "64(2)":{"d":"aggravated rape","m":9999},
        "66":{"d":"rape causing death or vegetative state","m":9999},
        "70":{"d":"gang rape","m":9999},
        "281":{"d":"rash driving on public way","m":6},
        "303":{"d":"theft","m":36},
        "304":{"d":"theft in dwelling house","m":84},
        "309":{"d":"robbery","m":120},
        "310":{"d":"dacoity","m":9999},
        "314":{"d":"criminal breach of trust","m":36},
        "316":{"d":"cheating","m":84},
        "318":{"d":"cheating and dishonestly inducing delivery of property","m":84},
        "85":{"d":"husband or relative subjecting woman to cruelty","m":36},
        "86":{"d":"dowry death","m":9999},
        "351":{"d":"criminal intimidation","m":24},
        "351(2)":{"d":"criminal intimidation — threat of death or grievous hurt","m":84},
        "352":{"d":"intentional insult to provoke breach of peace","m":24},
        "3(5)":{"d":"acts done in furtherance of common intention","m":0},
        "3(5)(a)":{"d":"acts done in furtherance of common intention","m":0},
        "190":{"d":"every member of unlawful assembly guilty of offence","m":0},
        "191":{"d":"rioting","m":24},
        "192":{"d":"rioting armed with deadly weapon","m":36},
    },
    "MV ACT":{
        "134":{"d":"failure to stop, render aid and report accident","m":6},
        "134(a)":{"d":"failure to stop vehicle and render aid after accident","m":6},
        "134(b)":{"d":"failure to report accident to police","m":6},
        "134(a)(b)":{"d":"failure to stop, render aid and report accident to police","m":6},
        "134(a)&(b)":{"d":"failure to stop, render aid and report accident to police","m":6},
        "184":{"d":"driving dangerously","m":12},
        "185":{"d":"driving under influence of alcohol or drugs","m":24},
        "187":{"d":"failure to comply with requirements after accident","m":3},
        "196":{"d":"driving without insurance","m":3},
    },
    "NDPS":{
        "8":{"d":"prohibition of certain operations","m":0},
        "20":{"d":"contravention in relation to cannabis","m":120},
        "20(b)(ii)":{"d":"possession of commercial quantity of cannabis","m":120},
        "21":{"d":"contravention in relation to manufactured drugs","m":120},
        "21(c)":{"d":"possession of commercial quantity of manufactured drugs","m":120},
        "22":{"d":"contravention in relation to psychotropic substances","m":120},
        "27":{"d":"consumption of narcotic drug or psychotropic substance","m":12},
        "27A":{"d":"financing illicit traffic and harbouring offenders","m":9999},
        "29":{"d":"abetment and criminal conspiracy","m":120},
    },
    "ARMS ACT":{
        "25":{"d":"illegal possession of arms or ammunition","m":36},
        "25(1A)":{"d":"illegal possession of prohibited arms","m":84},
        "27":{"d":"using arms in contravention","m":84},
        "29":{"d":"giving false information","m":6},
    },
    "SC/ST ACT":{
        "3":{"d":"atrocity against SC/ST person","m":60},
        "3(1)":{"d":"atrocity against SC/ST person","m":60},
        "3(1)(r)":{"d":"intentional insult or intimidation to SC/ST person","m":60},
        "3(1)(s)":{"d":"abusing SC/ST person by caste name in public","m":60},
        "3(2)":{"d":"offence against SC/ST by non-SC/ST person","m":9999},
        "3(2)(v)":{"d":"offence punishable with 10 years against SC/ST person","m":120},
        "4":{"d":"neglect of duties by public servant","m":12},
    },
    "DP ACT":{
        "3":{"d":"giving or taking dowry","m":60},
        "4":{"d":"demanding dowry","m":24},
    },
    "PC ACT":{
        "7":{"d":"public servant taking illegal gratification","m":84},
        "13":{"d":"criminal misconduct by public servant","m":120},
        "13(1)":{"d":"criminal misconduct by public servant","m":120},
        "15":{"d":"attempt to commit offence","m":36},
    },
    "POCSO":{
        "4":{"d":"penetrative sexual assault","m":120},
        "4(1)":{"d":"penetrative sexual assault","m":120},
        "5":{"d":"aggravated penetrative sexual assault","m":9999},
        "5(l)":{"d":"aggravated penetrative sexual assault on child below 16 years","m":9999},
        "5(k)":{"d":"aggravated penetrative sexual assault on child repeatedly","m":9999},
        "6":{"d":"aggravated penetrative sexual assault","m":9999},
        "6(1)":{"d":"aggravated penetrative sexual assault","m":9999},
        "7":{"d":"sexual assault","m":60},
        "8":{"d":"sexual assault","m":60},
        "9":{"d":"aggravated sexual assault","m":84},
        "10":{"d":"aggravated sexual assault","m":84},
        "11":{"d":"sexual harassment of child","m":36},
        "12":{"d":"sexual harassment of child","m":36},
    },
}

RW_ONLY = {"34","149","3(5)","3(5)(a)","3(5)(b)","190","49","16","17","18"}
ORDINALS = ["Firstly","Secondly","Thirdly","Fourthly","Fifthly",
            "Sixthly","Seventhly","Eighthly","Ninthly","Tenthly"]


# ═══════════════════════════════════════════════════════════════
# DATE HELPERS
# ═══════════════════════════════════════════════════════════════
def ord_sfx(n):
    if 11<=n<=13: return "th"
    return {1:"st",2:"nd",3:"rd"}.get(n%10,"th")

def legal_date(d=None):
    if d is None: d=date.today()
    M=["January","February","March","April","May","June",
       "July","August","September","October","November","December"]
    return f"{d.day}{ord_sfx(d.day)} day of {M[d.month-1]}, {d.year}"

def short_date(d=None):
    if d is None: d=date.today()
    return f"{d.day:02d}-{d.month:02d}-{d.year}"

# ═══════════════════════════════════════════════════════════════
# FILE HANDLING
# ═══════════════════════════════════════════════════════════════
def get_file_type(path):
    ext = Path(path).suffix.lower()
    if ext == '.docx': return 'docx'
    if ext == '.doc':  return 'doc'
    if ext == '.pdf':
        try:
            doc = fitz.open(path)
            text = "".join([p.get_text() for p in doc])
            doc.close()
            return 'scanned_pdf' if len(text.strip())<100 else 'typed_pdf'
        except: return 'scanned_pdf'
    return 'unknown'

def extract_pdf_text(path):
    doc = fitz.open(path)
    text = "".join([page.get_text()+"\n" for page in doc])
    doc.close()
    return text.strip()

def extract_docx_text(path):
    doc = DocxDoc(path)
    text = "\n".join([p.text for p in doc.paragraphs])
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells: text += "\n"+cell.text
    return text.strip()

def extract_doc_text(path):
    if HAS_MAMMOTH:
        try:
            with open(path,"rb") as f:
                return mammoth.extract_raw_text(f).value.strip()
        except: pass
    return ""

def pdf_to_images_b64(path):
    doc = fitz.open(path)
    images = []
    # Safety cap: only send the first MAX_SCANNED_PAGES pages to the AI.
    # Charge sheets run 3-4 pages typically; the upload screen also asks
    # for "charge sheet only" — this is a cost backstop, not the primary control.
    MAX_SCANNED_PAGES = 8
    for i, page in enumerate(doc):
        if i >= MAX_SCANNED_PAGES:
            break
        pix = page.get_pixmap(matrix=fitz.Matrix(2.0,2.0))
        images.append(base64.b64encode(pix.tobytes("png")).decode())
    doc.close()
    return images

def get_text(path, ftype, progress_cb=None):
    if ftype in ('typed_pdf','scanned_pdf'):
        if progress_cb: progress_cb("Reading PDF...")
        return extract_pdf_text(path)
    elif ftype == 'docx':
        if progress_cb: progress_cb("Reading Word document...")
        return extract_docx_text(path)
    elif ftype == 'doc':
        if progress_cb: progress_cb("Reading Word document (.doc)...")
        return extract_doc_text(path)
    return ""

# ═══════════════════════════════════════════════════════════════
# CLASSIFICATION
# ═══════════════════════════════════════════════════════════════
def lookup_section(sec, law):
    main_sec = sec.split(' r/w')[0].strip() if ' r/w' in sec else sec
    base = re.sub(r'\([^)]*\)','',main_sec).strip()
    for db_key in [law,"IPC","BNS","POCSO","MV ACT","NDPS","ARMS ACT","SC/ST ACT","DP ACT","PC ACT"]:
        db = DB.get(db_key,{})
        for v in [sec,main_sec,base,main_sec+"A"]:
            if v in db: return db[v], db_key
    return None, law

def classify_case(sections, law):
    max_m = 0; found = []
    for s in sections:
        if s.strip() in RW_ONLY: continue
        entry, act_key = lookup_section(s, law)
        if entry:
            found.append({"sec":s,"desc":entry["d"],"months":entry["m"],"act":act_key})
            if entry["m"] > max_m: max_m = entry["m"]
    return {"case_type":"WARRANT" if max_m>36 else "SUMMONS","max_m":max_m,"found":found}

def get_exam_ref(law, ct, court_type):
    if court_type == "sessions":
        return "228 Cr.P.C." if law=="IPC" else "240 BNSS"
    if law == "BNS":
        return "262 BNSS" if ct=="SUMMONS" else "274 BNSS"
    return "251 Cr.P.C." if ct=="SUMMONS" else "239 Cr.P.C."

def filter_substantive(sections):
    return [s for s in sections if s.strip() not in RW_ONLY and not s.strip().startswith('r/w')]

def extract_global_rw(sec_line):
    if not sec_line: return ""
    m = re.search(r'\br/w\.?\s+([\d\(\)A-Za-z\s\.]+?(?:IPC|BNS))\s*$',sec_line.strip(),re.I)
    return re.sub(r'\s+',' ',m.group(1)).strip() if m else ""

# ═══════════════════════════════════════════════════════════════
# MULTI-TURN AI FUNCTIONS
# ═══════════════════════════════════════════════════════════════

# ── TURN 1 — Extract all facts from charge sheet ──────────────
TURN1_PROMPT = """You are a senior legal analyst for Indian criminal courts in Andhra Pradesh.

READ ONLY the charge sheet portion. IGNORE completely:
- FIR copy, Memo of Evidence, List of Witnesses table, Inquest Report, PM Report, MVI Report, Arrest Memo
- Any content after "Hence the charge" line or after the police officer signature

The charge sheet starts with "IN THE COURT OF..." and ends with "Hence the charge."

Return ONLY a valid JSON object. No explanation. No markdown.

{
  "ps": "Police Station name from Sub-Inspector of Police [PS] P.S. or Sub-Divisional Police Officer [Place]",
  "io_designation": "Designation of Investigating Officer — Sub-Inspector of Police or Sub-Divisional Police Officer etc.",
  "accused": [
    {
      "name": "Full name exactly as written — include ALL initials and alias with @ symbol — do not drop any initial",
      "age": "Age as number string only — empty if not mentioned",
      "father": "Complete name after S/o or D/o or W/o — include ALL initials exactly as written — include the word Late if present",
      "village": "VILLAGE RULES below",
      "mandal": "MANDAL RULES below",
      "district": "Word(s) immediately before District — empty if not found",
      "full_addr": "Complete address of this accused from accused block",
      "exempted": false
    }
  ],
  "sec_line": "Exact section text from Charge sheet filed line — starting with U/Sec. or U/s. — include ALL sections and Act names exactly as written",
  "law": "IPC or BNS — primary law",
  "sections": ["SECTION RULES below"],
  "global_rw": "The r/w clause at end that applies to all sections — like 34 IPC or 3(5) BNS — empty if none",
  "cr_no": "Case number like 81/2025",
  "incident_date": "Date of incident — DD.MM.YYYY or period like August 2022 — empty if not found",
  "incident_time": "Time of incident — empty if not found",
  "incident_place": "Complete specific place of occurrence",
  "victim_name": "Name of deceased or victim",
  "victim_age": "Age of victim if mentioned",
  "accused_action": "What the accused specifically did — how the offence was committed",
  "injuries": "Injuries sustained by victim",
  "death_details": "How and where deceased died if applicable",
  "hospitals": "Hospitals shifted to if applicable",
  "pm_opinion": "Post Mortem or medical examination findings and doctor opinion — exact words — empty if not applicable",
  "mvi_opinion": "Motor Vehicle Inspector opinion if applicable — empty if not",
  "doctor_details": "Name and designation of doctor who conducted PM or medical examination",
  "mvi_details": "Name and designation of MVI if applicable",
  "witnesses": [
    {"lw_no": "LW.1", "name": "Full name of witness"}
  ],
  "mediators": [
    {"lw_no": "LW.12", "name": "Full name of mediator"}
  ],
  "doctors_in_lw": [
    {"lw_no": "LW.14", "name": "Name and designation of doctor"}
  ],
  "accused_fled": true,
  "case_type": "SUMMONS or WARRANT — based on maximum punishment — more than 3 years is WARRANT",
  "is_sessions": "true if triable by Sessions Court or Special Court"
}

SECTION EXTRACTION RULES — CRITICAL:
1. Split sections ONLY at commas and the word 'and' between COMPLETELY DIFFERENT sections
2. NEVER split at 'r/w' within a section — 'Sec.6 r/w 5(l)' is ONE section — extract as '6 r/w 5(l)'
3. The final r/w applying to ALL sections (like 'r/w 34 IPC' at the very end) goes in global_rw — NOT in sections array
4. Sub-clauses of same section are NOT separate sections — '498A' and '498A(1)' = ONE section only
5. Examples:
   'U/s 376(1) IPC and Sec.6 r/w 5(l) of POCSO Act 2012' → sections: ['376(1)', '6 r/w 5(l)']
   'U/s 118(1), 115(2), 351(2) r/w 3(5) BNS' → sections: ['118(1)', '115(2)', '351(2)'] global_rw: '3(5) BNS'
   'U/s 498A, 323 IPC r/w 34 IPC' → sections: ['498A', '323'] global_rw: '34 IPC'
   'U/s 106(1) BNS and Sec.134(a)(b) of MV Act 1988' → sections: ['106(1)', '134(a)(b)']

VILLAGE RULES (apply in order):
1. Word(s) immediately before Village keyword
2. Else word(s) before Town keyword
3. Else word(s) before Panchayat or Panchayath or Grama Panchayat
4. Else word(s) before H/w or H/wada or Hamlet
5. Else empty string

MANDAL RULES:
1. Word(s) immediately before Mandal keyword
2. Else word(s) before Taluqa or Taluka or Taluk or Taluq
3. Else word(s) before Town & Municipality — use entire phrase
4. Else word(s) before Town keyword
5. Else word(s) before City keyword
6. Else empty string — do NOT use village or district as fallback

WITNESS RULES:
- List ONLY witnesses examined by the SHO/IO from the charge sheet body (not from Memo of Evidence table)
- Witnesses who acted as mediators go in mediators array — not witnesses array
- Doctors and experts go in doctors_in_lw array — not witnesses array
- Format: LW.1, LW.2 etc.

IMPORTANT:
- List EVERY accused — count before answering
- Father name must include ALL initials — A. Subramanyam Reddy not just Subramanyam Reddy
- Return only the JSON — nothing else

CHARGE SHEET TEXT:
"""

# ── TURN 2 — Write Q.No.2 paragraph ──────────────────────────
def build_q2_prompt(facts, para_style, court_type, ps, io_desig):
    style_instruction = ""
    if para_style == "medium":
        style_instruction = """Write a MEDIUM length paragraph that covers all necessary points concisely.
Include: date time place, what accused did, injuries/death, hospital details, PM findings, doctor opinion, MVI opinion if applicable, accused fleeing if applicable, witnesses examined with LW numbers and names, mediators if any.
Do NOT include: detailed investigation procedure, arrest details, confession details, unnecessary background."""
    else:
        style_instruction = """Write an ELABORATED paragraph that covers every important point in detail.
Include: date time place, full description of how offence happened, vehicle details if any, victim details, injuries in detail, complete hospital journey, full PM findings with cause of death, doctor opinion in detail, MVI opinion if applicable, forensic findings if any, accused fleeing details, all witnesses examined with LW numbers and names, mediators if any, IO details.
Include every important fact from the charge sheet."""

    return f"""You are a legal document writer for Indian criminal courts.

Write the Q.No.2 brief paragraph for examination of accused based on these extracted facts:

{json.dumps(facts, indent=2, ensure_ascii=False)}

STRICT RULES:
1. Start with: "The {io_desig}, {ps} filed charge sheet against you that "
2. Use "accused of you [name]" — NEVER write "the accused" — always "accused of you"
3. After the facts — include doctor and MVI opinions naturally in the flow
4. After facts and opinions — include: "The {io_desig} examined [LW.1/Name, LW.2/Name...] and recorded their statements."
5. If mediators present — add: "[LW.X/Name and LW.Y/Name] acted as mediators at the time of arrest and recording of confession of accused of you."
6. End with: "What do you say?"
7. Write as ONE continuous paragraph — no line breaks
8. Use formal court language throughout

{style_instruction}

Write ONLY the paragraph — nothing else. No heading. No explanation."""

# ── TURN 3 — Write charge paragraphs ─────────────────────────
def build_charges_prompt(facts, sections, global_rw, court_desig, court_place):
    secs_text = ", ".join(sections)
    return f"""You are a legal document writer for Indian criminal courts.

Write the framing of charges paragraphs based on these facts:

{json.dumps(facts, indent=2, ensure_ascii=False)}

Sections to frame charges for: {secs_text}
Global r/w clause (applies to ALL sections): {global_rw if global_rw else 'None'}

STRICT FORMAT RULES:
1. Write one paragraph per section
2. Use ordinals: Firstly, Secondly, Thirdly... and ALWAYS use "Lastly" for the LAST charge
3. Each paragraph format: "[Ordinal]: That you on [date] at about [time], [place], [what accused specifically did for THIS section] and that you thereby committed an offence punishable under [Sec.X of Act] and [cognizance]."
4. First charge ends with: "within my cognizance"
5. All other charges end with: "within the cognizance of this Court"
6. If global r/w exists — attach it to EVERY section: "Sec.498A r/w 34 IPC" not just "Sec.498A IPC"
7. For r/w sections like "6 r/w 5(l)" — write: "Sec.6 r/w 5(l) of POCSO Act 2012"
8. Each charge paragraph must describe what accused did SPECIFICALLY for THAT section — different for each
9. Start verb phrase directly — NEVER write "The accused" — write "drove" or "committed" or "caused" etc.
10. Keep each charge short and specific — date, place, what accused did, section, cognizance
11. If only ONE section — use "Firstly" — no "Lastly" needed for single charge

Write ONLY the charge paragraphs — nothing else. No heading. No explanation. Separate each charge with a blank line."""

def call_claude(client, prompt, progress_cb=None, step=""):
    if progress_cb: progress_cb(f"AI thinking... {step}")
    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=3000,
        messages=[{"role":"user","content":prompt}]
    )
    return response.content[0].text.strip()

def ai_process(path, ftype, api_key, para_style, court_type, progress_cb=None):
    """Multi-turn AI processing — 3 focused turns for accuracy."""
    client = anthropic.Anthropic(api_key=api_key)

    # ── TURN 1 — Extract facts ────────────────────────────────
    if ftype == 'scanned_pdf':
        if progress_cb: progress_cb("Converting scanned PDF to images...")
        images = pdf_to_images_b64(path)
        if progress_cb: progress_cb("AI reading charge sheet... (Turn 1/3)")
        content = []
        for img in images:
            content.append({"type":"image","source":{"type":"base64",
                "media_type":"image/png","data":img}})
        content.append({"type":"text","text":TURN1_PROMPT.replace("CHARGE SHEET TEXT:","")})
        response = client.messages.create(
            model="claude-sonnet-4-5",
            max_tokens=3000,
            messages=[{"role":"user","content":content}]
        )
        raw = response.content[0].text.strip()
    else:
        text = get_text(path, ftype, progress_cb)
        if progress_cb: progress_cb("AI reading charge sheet... (Turn 1/3)")
        raw = call_claude(client, TURN1_PROMPT + text[:8000], progress_cb, "(Turn 1/3)")

    # Parse Turn 1 JSON
    raw = re.sub(r'^```json\s*','',raw,flags=re.M)
    raw = re.sub(r'^```\s*','',raw,flags=re.M)
    raw = re.sub(r'```\s*$','',raw,flags=re.M)
    raw = raw.strip()
    try: facts = json.loads(raw)
    except:
        m = re.search(r'\{[\s\S]+\}',raw)
        facts = json.loads(m.group()) if m else {}

    # Ensure fields
    defaults = {
        "ps":"","io_designation":"Sub-Inspector of Police","accused":[],
        "sec_line":"","law":"IPC","sections":[],"global_rw":"","cr_no":"",
        "incident_date":"","incident_time":"","incident_place":"","victim_name":"",
        "victim_age":"","accused_action":"","injuries":"","death_details":"",
        "hospitals":"","pm_opinion":"","mvi_opinion":"","doctor_details":"",
        "mvi_details":"","witnesses":[],"mediators":[],"doctors_in_lw":[],
        "accused_fled":False,"case_type":"","is_sessions":False
    }
    for k,v in defaults.items():
        if k not in facts: facts[k] = v
    if not facts["accused"]:
        facts["accused"] = [{"name":"","age":"","father":"","village":"",
            "mandal":"","district":"","full_addr":"","exempted":False}]
    for acc in facts["accused"]:
        if "exempted" not in acc: acc["exempted"] = False

    # ── TURN 2 — Write Q.No.2 paragraph ──────────────────────
    ps = facts.get("ps","")
    io_desig = facts.get("io_designation","Sub-Inspector of Police")
    q2_prompt = build_q2_prompt(facts, para_style, court_type, ps, io_desig)
    if progress_cb: progress_cb("AI writing Q.No.2 paragraph... (Turn 2/3)")
    q2_para = call_claude(client, q2_prompt, progress_cb, "(Turn 2/3)")
    facts["q2_para"] = q2_para.strip()

    # ── TURN 3 — Write charge paragraphs (if warrant/sessions) ─
    raw_secs = facts.get("sections",[])
    secs = filter_substantive(raw_secs)
    global_rw = facts.get("global_rw","")
    law = facts.get("law","IPC")

    db_anal = classify_case(raw_secs, law)
    db_ct   = db_anal["case_type"]
    ai_ct   = str(facts.get("case_type","")).upper().strip()
    if ai_ct not in ["SUMMONS","WARRANT"]: ai_ct = db_ct
    ct = "WARRANT" if (ai_ct=="WARRANT" or db_ct=="WARRANT") else "SUMMONS"
    ai_sessions = str(facts.get("is_sessions","")).lower() == "true"

    need_charges = (ct=="WARRANT" or court_type=="sessions" or ai_sessions) and secs

    facts["charge_paras"] = ""
    if need_charges:
        from docx import Document as _D
        court_cfg_court = ""  # will be passed from cfg
        charges_prompt = build_charges_prompt(facts, secs, global_rw, "", "")
        if progress_cb: progress_cb("AI writing framing of charges... (Turn 3/3)")
        charge_paras = call_claude(client, charges_prompt, progress_cb, "(Turn 3/3)")
        facts["charge_paras"] = charge_paras.strip()

    facts["case_type_resolved"] = ct
    facts["need_charges"] = need_charges
    facts["db_analysis"] = db_anal
    return facts

# ═══════════════════════════════════════════════════════════════
# DOCUMENT HELPERS
# ═══════════════════════════════════════════════════════════════
def build_heading(court):
    h = re.sub(r'\s+[Cc]ourt\s*$','',court.strip())
    h = re.sub(r'\s+[Cc]ourt\s*,',',',h)
    return f"IN THE COURT OF {h.upper()}"

def get_desig_parts(court):
    h = re.sub(r'\s+[Cc]ourt\s*$','',court.strip())
    h = re.sub(r'\s+[Cc]ourt\s*,',',',h)
    m = re.search(r'^(.+,)\s*([A-Z][a-z].+)\s*$',h)
    if m: return m.group(1).strip(), m.group(2).strip()
    return h,""

def set_cell_margins(cell, top=60, bottom=60, left=80, right=60):
    tc = cell._tc; tcPr = tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcMar')): tcPr.remove(old)
    tcMar = OxmlElement('w:tcMar')
    for side,val in [('top',top),('left',left),('bottom',bottom),('right',right)]:
        node = OxmlElement(f'w:{side}')
        node.set(qn('w:w'),str(val)); node.set(qn('w:type'),'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def fix_table_layout(table, col_widths_dxa):
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None: tblPr = OxmlElement('w:tblPr'); tbl.insert(0,tblPr)
    for tag in ['w:tblW','w:tblLayout']:
        old = tblPr.find(qn(tag))
        if old is not None: tblPr.remove(old)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'),str(sum(col_widths_dxa))); tblW.set(qn('w:type'),'dxa')
    tblPr.append(tblW)
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'),'fixed'); tblPr.append(tblLayout)
    tblGrid = tbl.find(qn('w:tblGrid'))
    if tblGrid is None: tblGrid = OxmlElement('w:tblGrid'); tbl.append(tblGrid)
    else:
        for c in list(tblGrid): tblGrid.remove(c)
    for w in col_widths_dxa:
        gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'),str(w)); tblGrid.append(gc)
    for row in table.rows:
        for i,cell in enumerate(row.cells):
            tc = cell._tc; tcPr = tc.get_or_add_tcPr()
            old = tcPr.find(qn('w:tcW'))
            if old is not None: tcPr.remove(old)
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'),str(col_widths_dxa[i])); tcW.set(qn('w:type'),'dxa')
            tcPr.append(tcW)

def add_run(para, text, bold=False, underline=False):
    run = para.add_run(text)
    run.font.name = "Arial"; run.font.size = Pt(11.5)
    run.bold = bold
    if underline: run.underline = True
    return run

def new_doc():
    from docx import Document as _Doc
    doc = _Doc()
    sec = doc.sections[0]
    sec.page_width=Cm(21.0); sec.page_height=Cm(29.7)
    sec.top_margin=Cm(2.54); sec.bottom_margin=Cm(2.54)
    sec.left_margin=Cm(2.54); sec.right_margin=Cm(2.54)
    return doc

def el(doc):
    p = doc.add_paragraph()
    p.paragraph_format.space_before=Pt(0)
    p.paragraph_format.space_after=Pt(0)
    add_run(p,''); return p

def add_desig_block(doc, d1, d2):
    TOTAL=9000; TW=5400; INDENT=TOTAL-TW
    dt=doc.add_table(rows=1,cols=1)
    dt_tbl=dt._tbl
    dt_tblPr=dt_tbl.find(qn('w:tblPr'))
    if dt_tblPr is None: dt_tblPr=OxmlElement('w:tblPr'); dt_tbl.insert(0,dt_tblPr)
    for tag in ['w:tblW','w:tblInd','w:tblBorders']:
        old=dt_tblPr.find(qn(tag))
        if old is not None: dt_tblPr.remove(old)
    tw_el=OxmlElement('w:tblW'); tw_el.set(qn('w:w'),str(TW)); tw_el.set(qn('w:type'),'dxa')
    dt_tblPr.append(tw_el)
    ti_el=OxmlElement('w:tblInd'); ti_el.set(qn('w:w'),str(INDENT)); ti_el.set(qn('w:type'),'dxa')
    dt_tblPr.append(ti_el)
    nb_el=OxmlElement('w:tblBorders')
    for edge in ['top','left','bottom','right','insideH','insideV']:
        e=OxmlElement(f'w:{edge}'); e.set(qn('w:val'),'none'); nb_el.append(e)
    dt_tblPr.append(nb_el)
    cell=dt.rows[0].cells[0]
    tc=cell._tc; tcPr=tc.get_or_add_tcPr()
    tcB=OxmlElement('w:tcBorders')
    for edge in ['top','left','bottom','right']:
        e=OxmlElement(f'w:{edge}'); e.set(qn('w:val'),'none'); tcB.append(e)
    tcPr.append(tcB)
    set_cell_margins(cell,top=0,bottom=0,left=0,right=0)
    p_d1=cell.paragraphs[0]; p_d1.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p_d1.paragraph_format.space_before=Pt(0); p_d1.paragraph_format.space_after=Pt(0)
    add_run(p_d1,d1,bold=True)
    if d2:
        p_d2=cell.add_paragraph(); p_d2.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p_d2.paragraph_format.space_before=Pt(0); p_d2.paragraph_format.space_after=Pt(0)
        add_run(p_d2,d2,bold=True)

# ═══════════════════════════════════════════════════════════════
# BUILD COMBINED WORD DOCUMENT
# ═══════════════════════════════════════════════════════════════
def build_document(facts, cfg, court_type, exam_ref, ct, manual_ref):
    court   = cfg.get("court","")
    judge   = cfg.get("judge","")
    heading = build_heading(court)
    d1,d2   = get_desig_parts(court)
    today   = date.today()
    sec_line= facts.get("sec_line","")
    q2_para = facts.get("q2_para","")
    charge_paras_text = facts.get("charge_paras","")
    need_charges = facts.get("need_charges",False)

    # Q.No.3
    if ct=="SUMMONS":
        q3 = f"The substance of accusation for the offence {sec_line}, read over and explained to you in vernacular language. Do you plead guilty or claim to be tried?"
    else:
        q3 = f"Charges {sec_line} have been framed against you, read over and explained to you in vernacular language. Do you plead guilty or claim to be tried?"

    doc  = new_doc()
    TOTAL=9000; DW=[2200,2600,1750,2450]; QW0=1050; QW1=TOTAL-QW0

    acc_list = [a for a in facts.get("accused",[{}]) if not a.get("exempted",False)]
    if not acc_list: acc_list = facts.get("accused",[{}])

    # ── EXAMINATION — one page per accused ──
    for idx,acc in enumerate(acc_list):
        if idx>0: doc.add_page_break()
        acc_no = f" No.{idx+1}" if len(acc_list)>1 else ""

        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(4)
        add_run(p,heading,bold=True,underline=True); el(doc)

        p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after=Pt(4)
        add_run(p2,f"Examination of the accused{acc_no} U/Sec.{exam_ref}",bold=True); el(doc)

        p3=doc.add_paragraph(); p3.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p3.paragraph_format.space_after=Pt(8)
        add_run(p3,"C.C. No._______________/_______________",bold=True); el(doc)

        det=doc.add_table(rows=4,cols=4); det.style='Table Grid'
        fix_table_layout(det,DW)
        rows_data=[
            ("Name :",acc.get("name",""),"District :",acc.get("district","")),
            ("Father's Name :",acc.get("father",""),"Calling :",""),
            ("Village :",acc.get("village",""),"Age :",f"{acc['age']} years" if acc.get("age") else ""),
            ("Mandal :",acc.get("mandal",""),"Date :",short_date(today)),
        ]
        for ri,(l1,v1,l2,v2) in enumerate(rows_data):
            row=det.rows[ri]
            for ci,(txt,bold) in enumerate([(l1,True),(v1,False),(l2,True),(v2,False)]):
                cell=row.cells[ci]
                set_cell_margins(cell,top=60,bottom=60,left=80,right=60)
                cp=cell.paragraphs[0]; cp.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
                add_run(cp,txt,bold=bold)
        el(doc)

        qa=[
            ("Q.No.1","Have you received the copies of documents in this case?"),
            ("Q.No.2",q2_para),("Q.No.3",q3),
            ("Q.No.4","Have you got means to engage an advocate?"),
        ]
        qt=doc.add_table(rows=len(qa)*2,cols=2); qt.style='Table Grid'
        fix_table_layout(qt,[QW0,QW1])
        ri=0
        for qno,qtext in qa:
            qrow=qt.rows[ri]
            set_cell_margins(qrow.cells[0],top=60,bottom=60,left=80,right=60)
            add_run(qrow.cells[0].paragraphs[0],qno,bold=True)
            set_cell_margins(qrow.cells[1],top=60,bottom=60,left=80,right=80)
            ap=qrow.cells[1].paragraphs[0]; ap.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            add_run(ap,qtext); ri+=1
            arow=qt.rows[ri]
            set_cell_margins(arow.cells[0],top=60,bottom=60,left=80,right=60)
            add_run(arow.cells[0].paragraphs[0],"Ans :",bold=True)
            set_cell_margins(arow.cells[1],top=60,bottom=60,left=80,right=80)
            for ln in range(2):
                p_line=arow.cells[1].paragraphs[0] if ln==0 else arow.cells[1].add_paragraph()
                p_line.paragraph_format.space_before=Pt(0)
                p_line.paragraph_format.space_after=Pt(0)
                p_line.paragraph_format.line_spacing=Pt(18)
                add_run(p_line,'')
            ri+=1

    # ── FRAMING OF CHARGES — next page ──
    if need_charges and charge_paras_text:
        # Page break — standalone paragraph
        body = doc.element.body
        sect_pr = body.find(qn('w:sectPr'))
        pb_para = OxmlElement('w:p')
        pb_run  = OxmlElement('w:r')
        pb_br   = OxmlElement('w:br')
        pb_br.set(qn('w:type'),'page')
        pb_run.append(pb_br); pb_para.append(pb_run)
        if sect_pr is not None:
            body.insert(list(body).index(sect_pr), pb_para)
        else:
            body.append(pb_para)

        # Court heading
        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after=Pt(4)
        add_run(p,heading,bold=True,underline=True); el(doc)

        p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after=Pt(8)
        add_run(p2,"C.C. No._______________/_______________",bold=True); el(doc)

        # Opening — Judge name + full designation + place
        if judge:
            opening = f"I, {judge}, {d1} {d2}, hereby charge you -"
        else:
            opening = f"I, Sri ___________________________, {d1} {d2}, hereby charge you -"
        p3=doc.add_paragraph(); p3.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        add_run(p3,opening); el(doc)

        # Accused box
        at=doc.add_table(rows=len(acc_list),cols=2); at.style='Table Grid'
        fix_table_layout(at,[600,8400])
        for idx,acc in enumerate(acc_list):
            prefix=f"A{idx+1}." if len(acc_list)>1 else ""
            full_addr=acc.get("full_addr","")
            if not full_addr:
                parts=[x for x in [acc.get("village",""),
                    (acc.get("mandal","")) if acc.get("mandal") else "",
                    (acc.get("district","")+" District") if acc.get("district") else ""] if x]
                full_addr=", ".join(parts)
            row=at.rows[idx]
            set_cell_margins(row.cells[0],top=60,bottom=60,left=80,right=60)
            add_run(row.cells[0].paragraphs[0],prefix,bold=True)
            set_cell_margins(row.cells[1],top=60,bottom=60,left=80,right=80)
            dp=row.cells[1].paragraphs[0]; dp.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            add_run(dp,acc.get("name",""),bold=True)
            age_s=f", aged {acc['age']} years," if acc.get("age") else ","
            father=acc.get("father","________________________")
            add_run(dp,f"{age_s} S/o. {father}, {full_addr}")
        el(doc)

        p4=doc.add_paragraph(); p4.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        add_run(p4,"as follows:"); el(doc)

        # Charge paragraphs from AI Turn 3
        charge_lines = [l.strip() for l in charge_paras_text.split('\n') if l.strip()]
        current_para = ""
        for line in charge_lines:
            if re.match(r'^(Firstly|Secondly|Thirdly|Fourthly|Fifthly|Sixthly|Seventhly|Lastly)\s*:',line):
                if current_para:
                    po=doc.add_paragraph(); po.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
                    po.paragraph_format.space_before=Pt(3); po.paragraph_format.space_after=Pt(3)
                    # Split ordinal from text
                    colon_idx = current_para.index(':')
                    ordinal_part = current_para[:colon_idx+1]
                    text_part = current_para[colon_idx+1:].strip()
                    r1=po.add_run(ordinal_part+" ")
                    r1.font.name="Arial"; r1.font.size=Pt(11.5); r1.bold=True
                    r2=po.add_run(text_part)
                    r2.font.name="Arial"; r2.font.size=Pt(11.5); r2.bold=False
                    el(doc)
                current_para = line
            else:
                current_para += " " + line if current_para else line

        # Last charge para
        if current_para:
            po=doc.add_paragraph(); po.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
            po.paragraph_format.space_before=Pt(3); po.paragraph_format.space_after=Pt(3)
            if ':' in current_para:
                colon_idx = current_para.index(':')
                ordinal_part = current_para[:colon_idx+1]
                text_part = current_para[colon_idx+1:].strip()
                r1=po.add_run(ordinal_part+" ")
                r1.font.name="Arial"; r1.font.size=Pt(11.5); r1.bold=True
                r2=po.add_run(text_part)
                r2.font.name="Arial"; r2.font.size=Pt(11.5); r2.bold=False
            else:
                add_run(po,current_para)
            el(doc)

        pd=doc.add_paragraph(); pd.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        add_run(pd,"And I, hereby direct that you be tried before me for the above said charges.")
        el(doc); el(doc)
        pd2=doc.add_paragraph(); pd2.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        add_run(pd2,f"Dated this the {legal_date(today)}.")
        el(doc); el(doc)
        add_desig_block(doc,d1,d2)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()

# ═══════════════════════════════════════════════════════════════
# MAIN PROCESSING PIPELINE
# ═══════════════════════════════════════════════════════════════
def process_file(path, cfg, court_type, manual_ref, para_style, progress_cb=None):
    api_key = cfg.get("api_key","")
    if not api_key: raise Exception("API key not set. Click Settings.")

    orig_name = Path(path).stem
    ftype     = get_file_type(path)

    # Multi-turn AI processing
    facts = ai_process(path, ftype, api_key, para_style, court_type, progress_cb)

    raw_secs = facts.get("sections",[])
    law      = facts.get("law","IPC")
    ct       = facts.get("case_type_resolved","SUMMONS")
    ai_sessions = str(facts.get("is_sessions","")).lower() == "true"

    # Determine exam ref
    if manual_ref:
        exam_ref = manual_ref
        ct = "SUMMONS" if any(x in manual_ref for x in ["251","262"]) else "WARRANT"
    elif court_type == "sessions" or ai_sessions:
        court_type = "sessions"
        exam_ref = "228 Cr.P.C." if law=="IPC" else "240 BNSS"
        ct = "WARRANT"
        facts["need_charges"] = True
    else:
        exam_ref = get_exam_ref(law, ct, court_type)

    # Remove exempted accused
    all_acc    = facts.get("accused",[{}])
    active_acc = [a for a in all_acc if not a.get("exempted",False)]
    if not active_acc: active_acc = all_acc
    facts["accused"] = active_acc

    if progress_cb: progress_cb("Generating Word document...")

    doc_bytes = build_document(facts, cfg, court_type, exam_ref, ct, manual_ref)

    return {
        "name":Path(path).name,"status":"ok","law":law,
        "sec_line":facts.get("sec_line",""),"sections":raw_secs,"ct":ct,
        "exam_ref":exam_ref,"ps":facts.get("ps",""),"cr_no":facts.get("cr_no",""),
        "acc_count":len(facts.get("accused",[])),"file_type":ftype,
        "docs":[{"fname":orig_name+".docx","data":doc_bytes}],
    }

# ═══════════════════════════════════════════════════════════════
# SELECTION DIALOG
# ═══════════════════════════════════════════════════════════════
