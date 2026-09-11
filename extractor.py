[extractor.py](https://github.com/user-attachments/files/32096107/extractor.py)
#!/usr/bin/env python3
"""
Charge Sheet Analyser — AI Powered
IPC/CrPC & BNS/BNSS | All Courts | Magistrate & Sessions
Powered by Anthropic Claude API
"""

import threading, os, sys, zipfile, re, json, io, base64
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
# SECTION DATABASE — All Acts
# ═══════════════════════════════════════════════════════════════
DB = {
    "IPC":{
        # Common sections
        "279":{"d":"Rash driving on public way","m":6},
        "283":{"d":"Danger or obstruction in public way","m":3},
        "302":{"d":"Murder","m":9999},"303":{"d":"Murder by life-convict","m":9999},
        "304":{"d":"Culpable homicide not amounting to murder","m":120},
        "304A":{"d":"Causing death by negligence","m":24},
        "304(A)":{"d":"Causing death by negligence","m":24},
        "304B":{"d":"Dowry death","m":9999},"304(B)":{"d":"Dowry death","m":9999},
        "305":{"d":"Abetment of suicide of child or insane person","m":120},
        "306":{"d":"Abetment of suicide","m":120},
        "307":{"d":"Attempt to murder","m":120},
        "308":{"d":"Attempt to commit culpable homicide","m":36},
        "309":{"d":"Attempt to commit suicide","m":12},
        "312":{"d":"Causing miscarriage","m":36},
        "313":{"d":"Causing miscarriage without consent","m":84},
        "319":{"d":"Hurt","m":0},"320":{"d":"Grievous hurt","m":0},
        "323":{"d":"Voluntarily causing hurt","m":12},
        "324":{"d":"Voluntarily causing hurt by dangerous weapons","m":36},
        "325":{"d":"Voluntarily causing grievous hurt","m":84},
        "326":{"d":"Voluntarily causing grievous hurt by dangerous weapons","m":9999},
        "326A":{"d":"Causing grievous hurt by use of acid","m":9999},
        "326B":{"d":"Voluntarily throwing acid","m":84},
        "341":{"d":"Wrongful restraint","m":1},
        "342":{"d":"Wrongful confinement","m":12},
        "352":{"d":"Assault or criminal force","m":3},
        "354":{"d":"Assault or criminal force to woman to outrage modesty","m":24},
        "354A":{"d":"Sexual harassment","m":36},
        "354B":{"d":"Assault with intent to disrobe","m":84},
        "354C":{"d":"Voyeurism","m":84},
        "354D":{"d":"Stalking","m":60},
        "363":{"d":"Kidnapping","m":84},
        "363A":{"d":"Kidnapping for begging","m":120},
        "364":{"d":"Kidnapping for murder","m":9999},
        "364A":{"d":"Kidnapping for ransom","m":9999},
        "365":{"d":"Kidnapping with intent to confine","m":84},
        "366":{"d":"Kidnapping woman to compel marriage","m":120},
        "376":{"d":"Rape","m":9999},
        "376A":{"d":"Rape causing death or vegetative state","m":9999},
        "376B":{"d":"Sexual intercourse by husband on separated wife","m":24},
        "376C":{"d":"Sexual intercourse by person in authority","m":120},
        "376D":{"d":"Gang rape","m":9999},
        "376DA":{"d":"Gang rape on woman under 16","m":9999},
        "376DB":{"d":"Gang rape on woman under 12","m":9999},
        "379":{"d":"Theft","m":36},
        "380":{"d":"Theft in dwelling house","m":84},
        "381":{"d":"Theft by clerk or servant","m":84},
        "382":{"d":"Theft with preparation for hurt","m":120},
        "392":{"d":"Robbery","m":120},
        "394":{"d":"Hurt in committing robbery","m":9999},
        "395":{"d":"Dacoity","m":9999},
        "396":{"d":"Dacoity with murder","m":9999},
        "397":{"d":"Robbery with attempt to cause death","m":84},
        "406":{"d":"Criminal breach of trust","m":36},
        "407":{"d":"Criminal breach of trust by carrier","m":84},
        "408":{"d":"Criminal breach of trust by clerk or servant","m":84},
        "409":{"d":"Criminal breach of trust by public servant","m":9999},
        "411":{"d":"Dishonestly receiving stolen property","m":36},
        "412":{"d":"Dishonestly receiving stolen property in gang robbery","m":9999},
        "417":{"d":"Cheating","m":12},
        "418":{"d":"Cheating with knowledge of wrongful loss","m":36},
        "419":{"d":"Cheating by personation","m":36},
        "420":{"d":"Cheating and dishonestly inducing delivery of property","m":84},
        "427":{"d":"Mischief causing damage","m":24},
        "435":{"d":"Mischief by fire to damage property","m":84},
        "436":{"d":"Mischief by fire to destroy house","m":9999},
        "447":{"d":"Criminal trespass","m":3},
        "448":{"d":"House-trespass","m":12},
        "452":{"d":"House-trespass after preparation for hurt","m":84},
        "454":{"d":"Lurking house-trespass","m":36},
        "457":{"d":"Lurking house-trespass by night","m":84},
        "458":{"d":"Lurking house-trespass with hurt","m":120},
        "465":{"d":"Forgery","m":24},
        "467":{"d":"Forgery of valuable security or will","m":9999},
        "468":{"d":"Forgery for purpose of cheating","m":84},
        "469":{"d":"Forgery for purpose of harming reputation","m":36},
        "471":{"d":"Using as genuine a forged document","m":24},
        "489A":{"d":"Counterfeiting currency notes","m":9999},
        "489B":{"d":"Using as genuine forged currency","m":9999},
        "498A":{"d":"Husband or relative subjecting woman to cruelty","m":36},
        "498(A)":{"d":"Husband or relative subjecting woman to cruelty","m":36},
        "504":{"d":"Intentional insult to provoke breach of peace","m":24},
        "506":{"d":"Criminal intimidation","m":24},
        "506(1)":{"d":"Criminal intimidation","m":24},
        "506(2)":{"d":"Criminal intimidation with death threat","m":84},
        "507":{"d":"Criminal intimidation by anonymous communication","m":24},
        "509":{"d":"Word or gesture to insult modesty of woman","m":3},
        "511":{"d":"Attempt to commit offences punishable with imprisonment","m":0},
        # Common intention / object
        "34":{"d":"Acts done in furtherance of common intention","m":0},
        "35":{"d":"Acts done with criminal knowledge or intention","m":0},
        "109":{"d":"Abetment if act is committed","m":0},
        "120B":{"d":"Criminal conspiracy","m":9999},
        "147":{"d":"Rioting","m":24},
        "148":{"d":"Rioting armed with deadly weapon","m":36},
        "149":{"d":"Every member of unlawful assembly guilty of offence","m":0},
        "186":{"d":"Obstructing public servant","m":3},
        "188":{"d":"Disobedience to order by public servant","m":6},
        "193":{"d":"False evidence","m":84},
        "201":{"d":"Causing disappearance of evidence of offence","m":84},
        "202":{"d":"Intentional omission to give information","m":6},
        "211":{"d":"False charge of offence made with intent to injure","m":84},
        "294":{"d":"Obscene acts and songs","m":3},
        "295A":{"d":"Deliberate acts outraging religious feelings","m":36},
        "420A":{"d":"Cheating by personation","m":36},
    },
    "BNS":{
        "103":{"d":"Murder","m":9999},
        "104":{"d":"Murder by life-convict","m":9999},
        "105":{"d":"Culpable homicide not amounting to murder","m":120},
        "106":{"d":"Causing death by negligence","m":60},
        "106(1)":{"d":"Causing death by negligence","m":60},
        "106(2)":{"d":"Causing death by negligence — hit and run","m":120},
        "107":{"d":"Causing death of quick unborn child","m":120},
        "108":{"d":"Abetment of suicide","m":120},
        "109":{"d":"Attempt to murder","m":120},
        "110":{"d":"Attempt to commit culpable homicide","m":36},
        "111":{"d":"Organised crime","m":9999},
        "113":{"d":"Terrorist act","m":9999},
        "115":{"d":"Voluntarily causing hurt","m":12},
        "115(2)":{"d":"Voluntarily causing grievous hurt","m":84},
        "116":{"d":"Voluntarily causing hurt by dangerous weapons","m":36},
        "117":{"d":"Voluntarily causing grievous hurt","m":84},
        "118":{"d":"Voluntarily causing grievous hurt by dangerous weapons","m":9999},
        "118(1)":{"d":"Voluntarily causing grievous hurt by dangerous weapons or means","m":9999},
        "118(2)":{"d":"Voluntarily causing grievous hurt by dangerous weapons causing death","m":9999},
        "119":{"d":"Voluntarily causing grievous hurt to extort property","m":120},
        "120":{"d":"Voluntarily causing hurt to extort confession","m":84},
        "121":{"d":"Assault or criminal force to woman to outrage modesty","m":24},
        "121(1)":{"d":"Assault or criminal force to woman to outrage modesty","m":24},
        "122":{"d":"Sexual harassment","m":36},
        "123":{"d":"Assault with intent to disrobe woman","m":84},
        "124":{"d":"Voyeurism","m":84},
        "125":{"d":"Stalking","m":60},
        "126":{"d":"Wrongful confinement","m":12},
        "127":{"d":"Wrongful restraint","m":1},
        "128":{"d":"Kidnapping","m":84},
        "130":{"d":"Kidnapping for murder","m":9999},
        "131":{"d":"Kidnapping for ransom","m":9999},
        "137":{"d":"Kidnapping woman to compel marriage","m":120},
        "64":{"d":"Rape","m":9999},
        "64(1)":{"d":"Rape","m":9999},
        "64(2)":{"d":"Aggravated rape","m":9999},
        "65":{"d":"Punishment for rape in certain cases","m":9999},
        "66":{"d":"Rape causing death or vegetative state","m":9999},
        "70":{"d":"Gang rape","m":9999},
        "281":{"d":"Rash driving on public way","m":6},
        "284":{"d":"Negligent conduct with respect to explosive","m":6},
        "303":{"d":"Theft","m":36},
        "304":{"d":"Theft in dwelling house","m":84},
        "309":{"d":"Robbery","m":120},
        "310":{"d":"Dacoity","m":9999},
        "311":{"d":"Dacoity with murder","m":9999},
        "314":{"d":"Criminal breach of trust","m":36},
        "316":{"d":"Cheating","m":84},
        "318":{"d":"Cheating and dishonestly inducing delivery of property","m":84},
        "319":{"d":"Cheating by personation","m":36},
        "324":{"d":"Mischief","m":36},
        "326":{"d":"Mischief by fire to destroy house","m":9999},
        "329":{"d":"Forgery","m":24},
        "334":{"d":"Forgery of valuable security","m":9999},
        "336":{"d":"Forgery for purpose of cheating","m":84},
        "338":{"d":"Forgery for purpose of harming reputation","m":36},
        "340":{"d":"Using as genuine a forged document","m":24},
        "356":{"d":"Defamation","m":24},
        "85":{"d":"Husband or relative subjecting woman to cruelty","m":36},
        "86":{"d":"Dowry death","m":9999},
        "351":{"d":"Criminal intimidation","m":24},
        "351(2)":{"d":"Criminal intimidation — threat of death or grievous hurt","m":84},
        "351(3)":{"d":"Criminal intimidation causing victim to do act","m":24},
        "352":{"d":"Intentional insult to provoke breach of peace","m":24},
        "353":{"d":"Statements conducing to public mischief","m":36},
        "152":{"d":"Acts endangering sovereignty of India","m":9999},
        "196":{"d":"Promoting enmity between groups","m":36},
        "3(5)":{"d":"Acts done in furtherance of common intention","m":0},
        "3(5)(a)":{"d":"Acts done in furtherance of common intention","m":0},
        "49":{"d":"Abetment if act is committed","m":0},
        "61":{"d":"Criminal conspiracy","m":9999},
        "191":{"d":"Rioting","m":24},
        "192":{"d":"Rioting armed with deadly weapon","m":36},
        "190":{"d":"Every member of unlawful assembly guilty of offence","m":0},
    },
    # Motor Vehicle Act 1988
    "MV ACT":{
        "134":{"d":"Duty of driver in case of accident and injury to persons","m":6},
        "134(a)":{"d":"Failure to stop vehicle after accident","m":6},
        "134(b)":{"d":"Failure to report accident to police","m":6},
        "134(a)(b)":{"d":"Failure to stop, render aid and report accident to police","m":6},
        "134(a)&(b)":{"d":"Failure to stop, render aid and report accident to police","m":6},
        "184":{"d":"Driving dangerously","m":12},
        "185":{"d":"Driving by drunken person or by person under influence of drugs","m":24},
        "187":{"d":"Failure to give information or comply with requirements","m":3},
        "196":{"d":"Driving without insurance","m":3},
        "197":{"d":"Taking vehicle without authority","m":3},
        "198":{"d":"Interference with vehicles","m":3},
        "199":{"d":"Offences relating to licensing of drivers","m":3},
    },
    # NDPS Act 1985
    "NDPS":{
        "8":{"d":"Prohibition of certain operations","m":0},
        "20":{"d":"Punishment for contravention in relation to cannabis","m":120},
        "20(a)":{"d":"Small quantity cannabis","m":6},
        "20(b)":{"d":"Cannabis — other than small quantity","m":120},
        "21":{"d":"Punishment for contravention in relation to manufactured drugs","m":120},
        "22":{"d":"Punishment for contravention in relation to psychotropic substances","m":120},
        "23":{"d":"Punishment for illegal import or export in India","m":120},
        "25":{"d":"Punishment for allowing premises to be used for drug offences","m":120},
        "27":{"d":"Punishment for illegal consumption of narcotic drugs","m":12},
        "27A":{"d":"Punishment for financing illicit traffic","m":9999},
        "28":{"d":"Punishment for attempts to commit offences","m":120},
        "29":{"d":"Punishment for abetment and criminal conspiracy","m":120},
        "35":{"d":"Presumption of culpable mental state","m":0},
    },
    # Arms Act 1959
    "ARMS ACT":{
        "25":{"d":"Punishment for certain offences — illegal possession of arms","m":36},
        "25(1)":{"d":"Illegal possession of firearms","m":36},
        "25(1A)":{"d":"Illegal possession of prohibited arms","m":84},
        "25(1B)":{"d":"Carrying firearms in public place","m":36},
        "26":{"d":"Punishment for contravention of licence conditions","m":36},
        "27":{"d":"Punishment for using arms in contravention","m":84},
        "29":{"d":"Giving false information","m":6},
        "30":{"d":"Punishment for contravention of provisions","m":36},
    },
    # SC/ST (Prevention of Atrocities) Act 1989
    "SC/ST ACT":{
        "3":{"d":"Punishment for offences of atrocities","m":60},
        "3(1)":{"d":"Various atrocities against SC/ST persons","m":60},
        "3(1)(r)":{"d":"Intentional insult or intimidation to SC/ST person","m":60},
        "3(1)(s)":{"d":"Abuses SC/ST person in public view","m":60},
        "3(2)":{"d":"False evidence to cause conviction of SC/ST person","m":9999},
        "3(2)(v)":{"d":"Offence against SC/ST person punishable with 10 years or more","m":120},
        "4":{"d":"Punishment for neglect of duties","m":12},
        "14":{"d":"Offence to be tried by Special Court","m":0},
    },
    # Dowry Prohibition Act 1961
    "DP ACT":{
        "3":{"d":"Penalty for giving or taking dowry","m":60},
        "4":{"d":"Penalty for demanding dowry","m":24},
        "4A":{"d":"Penalty for advertisements relating to dowry","m":6},
    },
    # Prevention of Corruption Act 1988
    "PC ACT":{
        "7":{"d":"Public servant taking gratification other than legal remuneration","m":84},
        "7A":{"d":"Taking undue advantage to influence public servant","m":84},
        "8":{"d":"Giving or taking bribe","m":84},
        "11":{"d":"Public servant obtaining valuable thing without consideration","m":36},
        "13":{"d":"Criminal misconduct by public servant","m":120},
        "13(1)":{"d":"Criminal misconduct by public servant","m":120},
        "15":{"d":"Punishment for attempt","m":36},
    },
    # POCSO Act 2012
    "POCSO":{
        "4":{"d":"Punishment for penetrative sexual assault","m":9999},
        "4(1)":{"d":"Penetrative sexual assault","m":120},
        "4(2)":{"d":"Aggravated penetrative sexual assault","m":9999},
        "5":{"d":"Aggravated penetrative sexual assault","m":9999},
        "6":{"d":"Punishment for aggravated penetrative sexual assault","m":9999},
        "7":{"d":"Sexual assault","m":60},
        "8":{"d":"Punishment for sexual assault","m":60},
        "9":{"d":"Aggravated sexual assault","m":84},
        "10":{"d":"Punishment for aggravated sexual assault","m":84},
        "11":{"d":"Sexual harassment of child","m":36},
        "12":{"d":"Punishment for sexual harassment","m":36},
        "13":{"d":"Use of child for pornographic purposes","m":60},
        "14":{"d":"Punishment for use of child for pornographic purposes","m":60},
        "17":{"d":"Punishment for abetment","m":0},
        "18":{"d":"Punishment for attempt to commit offence","m":0},
    },
}

# Sections that are never standalone charges
RW_ONLY = {"34","35","149","3(5)","3(5)(a)","3(5)(b)","109","49","190"}

ORDINALS = ["Firstly","Secondly","Thirdly","Fourthly","Fifthly",
            "Sixthly","Seventhly","Eighthly","Ninthly","Tenthly"]

# ═══════════════════════════════════════════════════════════════
# SETTINGS
# ═══════════════════════════════════════════════════════════════
SETTINGS_FILE = os.path.join(os.path.expanduser("~"),".csa_v2.json")

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
    if ftype == 'typed_pdf':
        if progress_cb: progress_cb("Reading PDF...")
        return extract_pdf_text(path)
    elif ftype == 'scanned_pdf':
        if progress_cb: progress_cb("Reading scanned PDF (OCR)...")
        return extract_pdf_text(path)
    elif ftype == 'docx':
        if progress_cb: progress_cb("Reading Word document...")
        return extract_docx_text(path)
    elif ftype == 'doc':
        if progress_cb: progress_cb("Reading Word document (.doc)...")
        return extract_doc_text(path)
    return ""

# ═══════════════════════════════════════════════════════════════
# AI EXTRACTION PROMPT
# ═══════════════════════════════════════════════════════════════
EXTRACT_PROMPT = """You are a senior legal analyst for Indian criminal courts.
Read ONLY the charge sheet portion of the document below.

IGNORE completely:
- Memo of Evidence / List of Witnesses
- FIR copy
- Inquest report
- Post Mortem report
- MVI report
- Arrest memo
- Any document after "Hence the charge" or after the Sub-Inspector signature

The charge sheet starts with the court heading "IN THE COURT OF..." and ends with "Hence the charge" or "Thus the accused...committed an offence...Hence the charge."

Extract ONLY from the charge sheet and return a valid JSON object. No explanation. No markdown. Just JSON.

{
  "ps": "Police Station name only from Sub-Inspector of Police [PS Name] P.S. line",
  "accused": [
    {
      "name": "Full name exactly as written including all initials and alias with @ symbol",
      "age": "Age as number string only",
      "father": "Complete father/mother name after S/o or D/o or W/o — include all initials exactly as written — keep the word Late or late if present",
      "village": "VILLAGE RULES — see below",
      "mandal": "MANDAL RULES — see below",
      "district": "Word(s) immediately before District keyword. Empty if not found.",
      "full_addr": "Complete address of this accused from the accused block",
      "exempted": false
    }
  ],
  "sec_line": "Exact section text from Charge sheet filed line — starting with U/Sec. or U/s. exactly as written up to and including the Act name. Example: U/s 106(1) BNS, Sec.134(a)(b) of Motor Vehicle Act-1988",
  "law": "IPC or BNS — primary law used",
  "sections": ["106(1)","134(a)(b)"],
  "acts": ["BNS","MV ACT"],
  "cr_no": "Case number like 81/2025",
  "incident_date": "Date of incident in DD.MM.YYYY format",
  "incident_time": "Time of incident like 10.00 PM or 4.30 p.m.",
  "incident_place": "Complete place of occurrence as described in charge sheet",
  "accused_act": "What the accused specifically did — one or two sentences — to be used in framing of charges",
  "victim_name": "Name of deceased or injured person if mentioned",
  "special_findings": "PM findings, MVI opinion, medical findings — one sentence if present, else empty",
  "escaped": "true if accused fled scene without aid or reporting, else false",
  "facts_brief": "Complete factual narrative covering ALL important points: date, time, place, what accused did, vehicle details if any, victim details, injuries sustained, hospital shifting, death details if applicable, PM findings, accused fleeing without aid if applicable, any special act violation. Cover every important point without limit. Use formal court English. No heading.",
  "case_type": "SUMMONS or WARRANT — based on maximum punishment. If any section carries more than 3 years it is WARRANT. Sections 34 IPC, 149 IPC, 3(5) BNS do not affect case type on their own.",
  "is_sessions": "true if offence is triable by Sessions Court (murder, rape, dacoity, etc.) else false"
}

VILLAGE RULES (apply in order for each accused):
1. Word(s) immediately before Village keyword
2. Else word(s) before Town keyword
3. Else word(s) before Panchayat or Panchayath or Grama Panchayat
4. Else word(s) before H/w or Hamlet
5. Else if none found — use same value as mandal

MANDAL RULES:
1. Word(s) immediately before Mandal keyword
2. Else word(s) before Taluqa or Taluka or Taluk or Taluk keyword (Taluk = Mandal)
3. Else word(s) before Town & Municipality — use entire "Town & Municipality" phrase
4. Else empty string

IMPORTANT:
- List EVERY accused found — count before answering
- If charge sheet has 3 accused the array must have 3 entries
- exempted: true if accused described as not charged, exonerated, no case made out
- facts_brief must cover ALL important facts — no artificial line limit
- Return only the JSON — nothing else

CHARGE SHEET TEXT:
"""

# ═══════════════════════════════════════════════════════════════
# AI EXTRACTION
# ═══════════════════════════════════════════════════════════════
def ai_extract(path, ftype, api_key, progress_cb=None):
    client = anthropic.Anthropic(api_key=api_key)

    if ftype == 'scanned_pdf':
        if progress_cb: progress_cb("Converting scanned PDF to images...")
        images = pdf_to_images_b64(path)
        if progress_cb: progress_cb("Sending to AI...")
        content = []
        for img in images:
            content.append({"type":"image","source":{"type":"base64",
                "media_type":"image/png","data":img}})
        content.append({"type":"text","text":EXTRACT_PROMPT.replace("CHARGE SHEET TEXT:","")})
    else:
        text = get_text(path, ftype, progress_cb)
        if progress_cb: progress_cb("Sending to AI...")
        content = [{"type":"text","text":EXTRACT_PROMPT + text[:8000]}]

    response = client.messages.create(
        model="claude-sonnet-4-5",
        max_tokens=3000,
        messages=[{"role":"user","content":content}]
    )

    raw = response.content[0].text.strip()
    raw = re.sub(r'^```json\s*','',raw,flags=re.M)
    raw = re.sub(r'^```\s*','',raw,flags=re.M)
    raw = re.sub(r'```\s*$','',raw,flags=re.M)
    raw = raw.strip()

    try: data = json.loads(raw)
    except:
        m = re.search(r'\{[\s\S]+\}',raw)
        if m:
            try: data = json.loads(m.group())
            except: data = {}
        else: data = {}

    # Ensure fields
    if "accused" not in data or not data["accused"]:
        data["accused"] = [{"name":"","age":"","father":"","village":"",
            "mandal":"","district":"","full_addr":"","exempted":False}]
    for acc in data["accused"]:
        if "exempted" not in acc: acc["exempted"] = False
    defaults = {"sections":[],"acts":[],"law":"IPC","sec_line":"","ps":"",
                "cr_no":"","incident_date":"","incident_time":"","incident_place":"",
                "accused_act":"","victim_name":"","special_findings":"","escaped":False,
                "facts_brief":"","case_type":"","is_sessions":False}
    for k,v in defaults.items():
        if k not in data: data[k] = v
    return data

# ═══════════════════════════════════════════════════════════════
# CASE CLASSIFICATION
# ═══════════════════════════════════════════════════════════════
def lookup_section(sec, law):
    """Look up a section in all databases."""
    base = re.sub(r'\([^)]*\)','',sec).strip()
    for db_key in [law, "MV ACT","NDPS","ARMS ACT","SC/ST ACT","DP ACT","PC ACT","POCSO"]:
        db = DB.get(db_key,{})
        for v in [sec, base, sec+"A", sec.replace("(","").replace(")","")]:
            if v in db: return db[v]
    return None

def classify_case(sections, law):
    max_m = 0; found = []
    for s in sections:
        entry = lookup_section(s, law)
        if entry:
            found.append({"sec":s,"desc":entry["d"],"months":entry["m"]})
            if entry["m"] > max_m: max_m = entry["m"]
    return {"case_type":"WARRANT" if max_m>36 else "SUMMONS",
            "max_m":max_m,"found":found}

def get_exam_ref(law, ct, court_type):
    """Get examination reference section."""
    if court_type == "sessions":
        return "228 Cr.P.C." if law=="IPC" else "240 BNSS"
    if law == "BNS":
        return "262 BNSS" if ct=="SUMMONS" else "274 BNSS"
    return "251 Cr.P.C." if ct=="SUMMONS" else "239 Cr.P.C."

def extract_rw_clause(sec_line):
    if not sec_line: return ""
    m = re.search(r'r/w\.?\s+([\d\(\)A-Za-z,\s&\.]+?(?:IPC|BNS|Act)[\-\d]*)\s*$',sec_line.strip(),re.I)
    return re.sub(r'\s+',' ',m.group(1)).strip() if m else ""

def filter_substantive(sections):
    return [s for s in sections if s.strip() not in RW_ONLY]

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
    return h, ""

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

def desig_block(doc, d1, d2):
    """Designation — borderless table pushed right, both lines centered."""
    TOTAL=9000; TW=5400; INDENT=TOTAL-TW
    NIL_VAL = {'style':'none'}
    dt = doc.add_table(rows=1,cols=1)
    dt_tbl = dt._tbl
    dt_tblPr = dt_tbl.find(qn('w:tblPr'))
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
    p_d1=cell.paragraphs[0]
    p_d1.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p_d1.paragraph_format.space_before=Pt(0); p_d1.paragraph_format.space_after=Pt(0)
    add_run(p_d1,d1,bold=True)
    if d2:
        p_d2=cell.add_paragraph()
        p_d2.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p_d2.paragraph_format.space_before=Pt(0); p_d2.paragraph_format.space_after=Pt(0)
        add_run(p_d2,d2,bold=True)

# ═══════════════════════════════════════════════════════════════
# EXAMINATION DOCUMENT
# ═══════════════════════════════════════════════════════════════
def make_examination_doc(info, cfg, court_type, exam_ref, ct):
    court   = cfg.get("court","")
    heading = build_heading(court)
    sec_line= info.get("sec_line","")
    facts   = info.get("facts_brief","") or "the facts as stated in the charge sheet"
    ps      = info.get("ps","")
    today   = date.today()

    io_str = f"Sub-Inspector of Police, {ps} P.S." if ps else "the Police"
    q2 = f"The {io_str} filed charge sheet against you that {facts} What do you say?"

    if ct == "SUMMONS":
        q3 = f"The substance of accusation for the offence {sec_line}, read over and explained to you in vernacular language. Do you plead guilty or claim to be tried?"
    else:
        q3 = f"Charges {sec_line} have been framed against you, read over and explained to you in vernacular language. Do you plead guilty or claim to be tried?"

    doc = new_doc()
    TOTAL=9000; DW=[2200,2600,1750,2450]; QW0=1050; QW1=TOTAL-QW0
    acc_list = [a for a in info.get("accused",[{}]) if not a.get("exempted",False)]
    if not acc_list: acc_list = info.get("accused",[{}])

    for idx,acc in enumerate(acc_list):
        if idx>0: doc.add_page_break()
        acc_no = f" No.{idx+1}" if len(acc_list)>1 else ""

        p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(4)
        add_run(p,heading,bold=True,underline=True)
        el(doc)

        p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_after=Pt(4)
        add_run(p2,f"Examination of the accused{acc_no} U/Sec.{exam_ref}",bold=True)
        el(doc)

        p3=doc.add_paragraph(); p3.alignment=WD_ALIGN_PARAGRAPH.CENTER
        p3.paragraph_format.space_after=Pt(8)
        add_run(p3,"C.C. No._______________/_______________",bold=True)
        el(doc)

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
            ("Q.No.2",q2),("Q.No.3",q3),
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

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()

# ═══════════════════════════════════════════════════════════════
# FRAMING OF CHARGES DOCUMENT
# ═══════════════════════════════════════════════════════════════
def make_charges_doc(info, cfg, court_type, law):
    court   = cfg.get("court","")
    judge   = cfg.get("judge","")
    heading = build_heading(court)
    d1,d2   = get_desig_parts(court)
    today   = date.today()
    sec_line= info.get("sec_line","")
    raw_secs= info.get("sections",[])
    secs    = filter_substantive(raw_secs)
    rw      = extract_rw_clause(sec_line)
    analysis= classify_case(raw_secs,law)
    idate   = info.get("incident_date","")
    itime   = info.get("incident_time","")
    iplace  = info.get("incident_place","")
    acc_act = info.get("accused_act","")
    victim  = info.get("victim_name","")
    special = info.get("special_findings","")
    escaped = info.get("escaped",False)

    judge_str = judge if judge else f"Sri ___________________________, {d1} {d2}"

    doc = new_doc()
    TOTAL=9000

    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_after=Pt(4)
    add_run(p,heading,bold=True,underline=True); el(doc)

    p2=doc.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p2.paragraph_format.space_after=Pt(8)
    add_run(p2,"C.C. No._______________/_______________",bold=True); el(doc)

    p3=doc.add_paragraph(); p3.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(p3,f"I, {judge_str}, hereby charge you -"); el(doc)

    acc_list=[a for a in info.get("accused",[{}]) if not a.get("exempted",False)]
    if not acc_list: acc_list=info.get("accused",[{}])

    at=doc.add_table(rows=len(acc_list),cols=2); at.style='Table Grid'
    fix_table_layout(at,[600,8400])
    for idx,acc in enumerate(acc_list):
        prefix=f"A{idx+1}." if len(acc_list)>1 else ""
        full_addr=acc.get("full_addr","")
        if not full_addr:
            parts=[x for x in [acc.get("village",""),
                (acc.get("mandal","")+" Mandal") if acc.get("mandal") else "",
                (acc.get("district","")+" District") if acc.get("district") else ""] if x]
            full_addr=", ".join(parts)
        row=at.rows[idx]
        set_cell_margins(row.cells[0],top=60,bottom=60,left=80,right=60)
        add_run(row.cells[0].paragraphs[0],prefix,bold=True)
        set_cell_margins(row.cells[1],top=60,bottom=60,left=80,right=80)
        dp=row.cells[1].paragraphs[0]; dp.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        add_run(dp,acc.get("name",""),bold=True)
        age_s=f", aged {acc['age']} years," if acc.get("age") else ","
        add_run(dp,f"{age_s} S/o. {acc.get('father','________________________')}, {full_addr}")
    el(doc)

    p4=doc.add_paragraph(); p4.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(p4,"as follows:"); el(doc)

    # Build charge paragraphs
    all_secs=[]
    for s in secs:
        e=next((x for x in analysis["found"] if x["sec"]==s),None)
        all_secs.append(e if e else {"sec":s,"desc":"the offence as stated in the charge sheet","months":0})

    date_time_place = ""
    if idate and itime and iplace:
        date_time_place = f"on {idate} at about {itime}, {iplace}"
    elif idate and iplace:
        date_time_place = f"on {idate}, {iplace}"
    else:
        date_time_place = "on the date, time and place mentioned in the charge sheet"

    total = len(all_secs)
    for idx,se in enumerate(all_secs):
        is_last = (idx == total-1)
        is_first = (idx == 0)
        ordinal = "Lastly" if is_last and total>1 else (ORDINALS[idx] if idx<len(ORDINALS) else f"{idx+1}thly")
        cognizance = "within my cognizance" if is_first else "within the cognizance of this Court"
        sec_with_rw = f"{se['sec']} r/w {rw}" if rw else f"Sec.{se['sec']} of {law}"

        # Build short charge text
        if acc_act:
            charge_body = f"{acc_act}"
        elif victim:
            charge_body = f"caused {se['desc']} to {victim}"
        else:
            charge_body = f"committed the offence of {se['desc']}"

        charge_text = f"That you {date_time_place}, {charge_body} and that you thereby committed an offence punishable under {sec_with_rw} and {cognizance}."

        # Ordinal + charge text in one justified paragraph
        po=doc.add_paragraph(); po.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
        po.paragraph_format.space_before=Pt(3); po.paragraph_format.space_after=Pt(3)
        r1=po.add_run(f"{ordinal}: ")
        r1.font.name="Arial"; r1.font.size=Pt(11.5); r1.bold=True
        r2=po.add_run(charge_text)
        r2.font.name="Arial"; r2.font.size=Pt(11.5); r2.bold=False
        el(doc)

    pd=doc.add_paragraph(); pd.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(pd,"And I, hereby direct that you be tried before me for the above said charges.")
    el(doc); el(doc)

    pd2=doc.add_paragraph(); pd2.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    add_run(pd2,f"Dated this the {legal_date(today)}.")
    el(doc); el(doc)

    desig_block(doc,d1,d2)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()

# ═══════════════════════════════════════════════════════════════
# PROCESSING PIPELINE
# ═══════════════════════════════════════════════════════════════
def process_file(path, cfg, court_type, manual_ref, progress_cb=None):
    api_key = cfg.get("api_key","")
    if not api_key:
        raise Exception("API key not set. Click Settings.")

    base  = Path(path).stem
    ftype = get_file_type(path)

    info = ai_extract(path, ftype, api_key, progress_cb)

    raw_secs = info.get("sections",[])
    law      = info.get("law","IPC")
    db_anal  = classify_case(raw_secs, law)
    db_ct    = db_anal["case_type"]
    ai_ct    = str(info.get("case_type","")).upper().strip()
    if ai_ct not in ["SUMMONS","WARRANT"]: ai_ct = db_ct
    ai_sessions = str(info.get("is_sessions","")).lower() == "true"

    # Determine exam ref and case type
    if manual_ref:
        exam_ref = manual_ref
        if "251" in manual_ref or "262" in manual_ref: ct = "SUMMONS"
        else: ct = "WARRANT"
    elif court_type == "sessions":
        exam_ref = "228 Cr.P.C." if law=="IPC" else "240 BNSS"
        ct = "WARRANT"
    else:
        ct = "WARRANT" if (ai_ct=="WARRANT" or db_ct=="WARRANT") else "SUMMONS"
        exam_ref = get_exam_ref(law, ct, court_type)

    info["case_type_resolved"] = ct

    # Remove exempted
    all_acc    = info.get("accused",[{}])
    active_acc = [a for a in all_acc if not a.get("exempted",False)]
    if not active_acc: active_acc = all_acc
    info["accused"] = active_acc

    if progress_cb: progress_cb("Generating Word document...")

    doc = new_doc()
    doc_bytes_list = []

    exam_bytes   = make_examination_doc(info, cfg, court_type, exam_ref, ct)
    need_charges = (ct=="WARRANT" or court_type=="sessions") and filter_substantive(raw_secs)

    docs = []
    docs.append({"fname":base+"_Examination.docx","data":exam_bytes})
    if need_charges:
        charges_bytes = make_charges_doc(info, cfg, court_type, law)
        docs.append({"fname":base+"_Framing_of_Charges.docx","data":charges_bytes})

    return {
        "name":Path(path).name,"status":"ok","law":law,
        "sec_line":info.get("sec_line",""),"sections":raw_secs,"ct":ct,
        "exam_ref":exam_ref,"ps":info.get("ps",""),"cr_no":info.get("cr_no",""),
        "acc_count":len(info.get("accused",[])),"file_type":ftype,"docs":docs,
    }

# ═══════════════════════════════════════════════════════════════
# COURT TYPE + MANUAL SELECTION DIALOG
# ═══════════════════════════════════════════════════════════════
