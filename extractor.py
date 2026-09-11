#!/usr/bin/env python3
"""
Charge Sheet Analyser — AI Powered
IPC/CrPC & BNS/BNSS | All Courts | Magistrate & Sessions
Powered by Anthropic Claude API
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
# COMPLETE SECTION DATABASE
# ═══════════════════════════════════════════════════════════════
DB = {
    "IPC":{
        "279":{"d":"rash driving or riding on a public way","m":6},
        "283":{"d":"causing danger or obstruction in public way","m":3},
        "285":{"d":"negligent conduct with respect to fire or combustible matter","m":6},
        "286":{"d":"negligent conduct with respect to explosive substance","m":6},
        "287":{"d":"negligent conduct with respect to machinery","m":6},
        "289":{"d":"negligent conduct with respect to animal","m":6},
        "302":{"d":"murder","m":9999},
        "303":{"d":"murder by life-convict","m":9999},
        "304":{"d":"culpable homicide not amounting to murder","m":120},
        "304A":{"d":"causing death by negligence","m":24},
        "304(A)":{"d":"causing death by negligence","m":24},
        "304B":{"d":"dowry death","m":9999},
        "304(B)":{"d":"dowry death","m":9999},
        "305":{"d":"abetment of suicide of child or insane person","m":120},
        "306":{"d":"abetment of suicide","m":120},
        "307":{"d":"attempt to murder","m":120},
        "308":{"d":"attempt to commit culpable homicide","m":36},
        "309":{"d":"attempt to commit suicide","m":12},
        "312":{"d":"causing miscarriage","m":36},
        "313":{"d":"causing miscarriage without woman's consent","m":84},
        "315":{"d":"act done with intent to prevent child being born alive","m":120},
        "319":{"d":"hurt","m":0},
        "320":{"d":"grievous hurt","m":0},
        "323":{"d":"voluntarily causing hurt","m":12},
        "324":{"d":"voluntarily causing hurt by dangerous weapons or means","m":36},
        "325":{"d":"voluntarily causing grievous hurt","m":84},
        "326":{"d":"voluntarily causing grievous hurt by dangerous weapons or means","m":9999},
        "326A":{"d":"voluntarily causing grievous hurt by use of acid","m":9999},
        "326B":{"d":"voluntarily throwing or attempting to throw acid","m":84},
        "341":{"d":"wrongful restraint","m":1},
        "342":{"d":"wrongful confinement","m":12},
        "343":{"d":"wrongful confinement for three or more days","m":24},
        "344":{"d":"wrongful confinement for ten or more days","m":36},
        "345":{"d":"wrongful confinement of person for whose liberation writ has been issued","m":24},
        "346":{"d":"wrongful confinement in secret","m":24},
        "347":{"d":"wrongful confinement to extort property","m":36},
        "348":{"d":"wrongful confinement to extort confession","m":36},
        "352":{"d":"assault or criminal force otherwise than on grave provocation","m":3},
        "353":{"d":"assault or criminal force to deter public servant from discharge of his duty","m":24},
        "354":{"d":"assault or criminal force to woman with intent to outrage her modesty","m":24},
        "354A":{"d":"sexual harassment","m":36},
        "354B":{"d":"assault or use of criminal force to woman with intent to disrobe","m":84},
        "354C":{"d":"voyeurism","m":84},
        "354D":{"d":"stalking","m":60},
        "355":{"d":"assault or criminal force with intent to dishonour person","m":24},
        "363":{"d":"kidnapping","m":84},
        "363A":{"d":"kidnapping or maiming a minor for purposes of begging","m":120},
        "364":{"d":"kidnapping or abducting in order to murder","m":9999},
        "364A":{"d":"kidnapping for ransom","m":9999},
        "365":{"d":"kidnapping or abducting with intent to confine person","m":84},
        "366":{"d":"kidnapping abducting or inducing woman to compel her marriage","m":120},
        "366A":{"d":"procuration of minor girl","m":120},
        "366B":{"d":"importation of girl from foreign country","m":120},
        "367":{"d":"kidnapping or abducting in order to subject person to grievous hurt slavery","m":120},
        "368":{"d":"wrongfully concealing or keeping in confinement kidnapped person","m":84},
        "370":{"d":"trafficking of person","m":84},
        "370A":{"d":"exploitation of a trafficked person","m":84},
        "371":{"d":"habitual dealing in slaves","m":9999},
        "372":{"d":"selling minor for purposes of prostitution","m":120},
        "373":{"d":"buying minor for purposes of prostitution","m":120},
        "374":{"d":"unlawful compulsory labour","m":12},
        "375":{"d":"rape","m":9999},
        "376":{"d":"rape","m":9999},
        "376(1)":{"d":"rape","m":9999},
        "376(2)":{"d":"aggravated rape","m":9999},
        "376A":{"d":"rape causing death or vegetative state","m":9999},
        "376AB":{"d":"rape on woman under twelve years of age","m":9999},
        "376B":{"d":"sexual intercourse by husband upon his wife during separation","m":24},
        "376C":{"d":"sexual intercourse by person in authority","m":120},
        "376D":{"d":"gang rape","m":9999},
        "376DA":{"d":"gang rape on woman under sixteen years of age","m":9999},
        "376DB":{"d":"gang rape on woman under twelve years of age","m":9999},
        "376E":{"d":"punishment for repeat offenders","m":9999},
        "377":{"d":"unnatural offences","m":9999},
        "379":{"d":"theft","m":36},
        "380":{"d":"theft in dwelling house","m":84},
        "381":{"d":"theft by clerk or servant of property in possession of master","m":84},
        "382":{"d":"theft after preparation made for causing death hurt or restraint in order to the committing of the theft","m":120},
        "384":{"d":"extortion","m":36},
        "385":{"d":"putting person in fear of injury in order to commit extortion","m":24},
        "386":{"d":"extortion by putting a person in fear of death or grievous hurt","m":120},
        "387":{"d":"putting person in fear of death or of grievous hurt in order to commit extortion","m":84},
        "388":{"d":"extortion by threat of accusation of an offence punishable with death or imprisonment for life","m":120},
        "389":{"d":"putting person in fear of accusation of offence in order to commit extortion","m":84},
        "392":{"d":"robbery","m":120},
        "393":{"d":"attempt to commit robbery","m":84},
        "394":{"d":"voluntarily causing hurt in committing robbery","m":9999},
        "395":{"d":"dacoity","m":9999},
        "396":{"d":"dacoity with murder","m":9999},
        "397":{"d":"robbery or dacoity with attempt to cause death or grievous hurt","m":84},
        "398":{"d":"attempt to commit robbery or dacoity when armed with deadly weapon","m":84},
        "399":{"d":"making preparation to commit dacoity","m":120},
        "400":{"d":"punishment of member of gang of dacoits","m":9999},
        "401":{"d":"punishment of belonging to gang of thieves","m":84},
        "402":{"d":"assembling for purpose of committing dacoity","m":84},
        "403":{"d":"dishonest misappropriation of property","m":24},
        "404":{"d":"dishonest misappropriation of property possessed by deceased person at the time of his death","m":36},
        "405":{"d":"criminal breach of trust","m":0},
        "406":{"d":"punishment for criminal breach of trust","m":36},
        "407":{"d":"criminal breach of trust by carrier merchant or agent","m":84},
        "408":{"d":"criminal breach of trust by clerk or servant","m":84},
        "409":{"d":"criminal breach of trust by public servant or by banker merchant or agent","m":9999},
        "410":{"d":"stolen property","m":0},
        "411":{"d":"dishonestly receiving stolen property","m":36},
        "412":{"d":"dishonestly receiving property stolen in the commission of a dacoity","m":9999},
        "413":{"d":"habitually dealing in stolen property","m":9999},
        "414":{"d":"assisting in concealment of stolen property","m":36},
        "415":{"d":"cheating","m":0},
        "416":{"d":"cheating by personation","m":0},
        "417":{"d":"punishment for cheating","m":12},
        "418":{"d":"cheating with knowledge that wrongful loss may ensue to person whose interest offender is bound to protect","m":36},
        "419":{"d":"punishment for cheating by personation","m":36},
        "420":{"d":"cheating and dishonestly inducing delivery of property","m":84},
        "421":{"d":"dishonest or fraudulent removal or concealment of property to prevent distribution among creditors","m":24},
        "424":{"d":"dishonest or fraudulent execution of deed of transfer containing false statement of consideration","m":24},
        "425":{"d":"mischief","m":0},
        "426":{"d":"punishment for mischief","m":3},
        "427":{"d":"mischief causing damage to the amount of fifty rupees","m":24},
        "428":{"d":"mischief by killing or maiming animal of the value of ten rupees","m":24},
        "429":{"d":"mischief by killing or maiming cattle etc. of any value or any animal of the value of fifty rupees","m":60},
        "430":{"d":"mischief by injury to works of irrigation or by wrongfully diverting water","m":60},
        "431":{"d":"mischief by injury to public road bridge river or channel","m":60},
        "432":{"d":"mischief by causing inundation or obstruction to public drainage attended with damage","m":60},
        "433":{"d":"mischief by destroying moving or rendering less useful a light-house or sea-mark","m":84},
        "434":{"d":"mischief by destroying or moving etc. a land-mark fixed by public authority","m":12},
        "435":{"d":"mischief by fire or explosive substance with intent to cause damage to amount of one hundred rupees","m":84},
        "436":{"d":"mischief by fire or explosive substance with intent to destroy house etc.","m":9999},
        "437":{"d":"mischief with intent to destroy or make unsafe a decked vessel or one of twenty tons burden","m":120},
        "438":{"d":"punishment of mischief committed after preparation made for causing death or hurt","m":120},
        "439":{"d":"punishment for intentionally running vessel aground or ashore with intent to commit theft etc.","m":120},
        "440":{"d":"mischief committed after preparation made for causing death or hurt","m":60},
        "441":{"d":"criminal trespass","m":0},
        "447":{"d":"punishment for criminal trespass","m":3},
        "448":{"d":"punishment for house-trespass","m":12},
        "449":{"d":"house-trespass in order to commit offence punishable with death","m":9999},
        "450":{"d":"house-trespass in order to commit offence punishable with imprisonment for life","m":9999},
        "451":{"d":"house-trespass in order to commit offence punishable with imprisonment","m":24},
        "452":{"d":"house-trespass after preparation for hurt assault or wrongful restraint","m":84},
        "453":{"d":"punishment for lurking house-trespass or house-breaking","m":24},
        "454":{"d":"lurking house-trespass or house-breaking in order to commit offence punishable with imprisonment","m":36},
        "455":{"d":"lurking house-trespass or house-breaking after preparation for hurt assault or wrongful restraint","m":120},
        "456":{"d":"punishment for lurking house-trespass or house-breaking by night","m":36},
        "457":{"d":"lurking house-trespass or house-breaking by night in order to commit offence punishable with imprisonment","m":84},
        "458":{"d":"lurking house-trespass or house-breaking by night after preparation for hurt assault or wrongful restraint","m":120},
        "459":{"d":"grievous hurt caused whilst committing lurking house-trespass or house-breaking","m":9999},
        "460":{"d":"all persons jointly concerned in lurking house-trespass or house-breaking by night punishable where death or grievous hurt caused by one of them","m":9999},
        "461":{"d":"dishonestly breaking open receptacle containing property","m":24},
        "462":{"d":"punishment for same offence when committed by person entrusted with custody","m":36},
        "463":{"d":"forgery","m":0},
        "465":{"d":"punishment for forgery","m":24},
        "466":{"d":"forgery of record of Court or of public register etc.","m":84},
        "467":{"d":"forgery of valuable security will etc.","m":9999},
        "468":{"d":"forgery for purpose of cheating","m":84},
        "469":{"d":"forgery for purpose of harming reputation","m":36},
        "470":{"d":"forged document","m":0},
        "471":{"d":"using as genuine a forged document or electronic record","m":24},
        "472":{"d":"making or possessing counterfeit seal etc. with intent to commit forgery punishable under section 467","m":9999},
        "473":{"d":"making or possessing counterfeit seal etc. with intent to commit forgery punishable otherwise","m":36},
        "474":{"d":"having possession of document described in section 466 or 467 knowing it to be forged and intending to use it as genuine","m":9999},
        "475":{"d":"counterfeiting device or mark used for authenticating documents described in section 467 or possessing counterfeit marked material","m":9999},
        "476":{"d":"counterfeiting device or mark used for authenticating documents other than those described in section 467","m":84},
        "477":{"d":"fraudulent cancellation destruction etc. of will authority to adopt or valuable security","m":9999},
        "477A":{"d":"falsification of accounts","m":84},
        "489A":{"d":"counterfeiting currency-notes or bank-notes","m":9999},
        "489B":{"d":"using as genuine forged or counterfeit currency-notes or bank-notes","m":9999},
        "489C":{"d":"possession of forged or counterfeit currency-notes or bank-notes","m":84},
        "489D":{"d":"making or possessing instruments or materials for forging or counterfeiting currency-notes or bank-notes","m":9999},
        "489E":{"d":"making use of documents resembling currency-notes or bank-notes","m":12},
        "491":{"d":"breach of contract to attend on and supply wants of helpless person","m":36},
        "493":{"d":"cohabitation caused by a man deceitfully inducing a belief of lawful marriage","m":120},
        "494":{"d":"marrying again during lifetime of husband or wife","m":84},
        "495":{"d":"same offence with concealment of former marriage from person with whom subsequent marriage is contracted","m":120},
        "496":{"d":"marriage ceremony fraudulently gone through without lawful marriage","m":84},
        "497":{"d":"adultery","m":60},
        "498":{"d":"enticing or taking away or detaining with criminal intent a married woman","m":24},
        "498A":{"d":"husband or relative of husband subjecting woman to cruelty","m":36},
        "498(A)":{"d":"husband or relative of husband subjecting woman to cruelty","m":36},
        "499":{"d":"defamation","m":0},
        "500":{"d":"punishment for defamation","m":24},
        "501":{"d":"printing or engraving matter known to be defamatory","m":24},
        "502":{"d":"sale of printed or engraved substance containing defamatory matter","m":24},
        "503":{"d":"criminal intimidation","m":0},
        "504":{"d":"intentional insult with intent to provoke breach of the peace","m":24},
        "505":{"d":"statements conducing to public mischief","m":36},
        "506":{"d":"punishment for criminal intimidation","m":24},
        "506(1)":{"d":"criminal intimidation","m":24},
        "506(2)":{"d":"criminal intimidation with threat of death or grievous hurt","m":84},
        "507":{"d":"criminal intimidation by anonymous communication","m":24},
        "508":{"d":"act caused by inducing person to believe that he will be rendered an object of the divine displeasure","m":12},
        "509":{"d":"word gesture or act intended to insult the modesty of a woman","m":36},
        "510":{"d":"misconduct in public by a drunken person","m":1},
        "511":{"d":"punishment for attempting to commit offences punishable with imprisonment for life or other imprisonment","m":0},
        # Common intention / abetment
        "34":{"d":"acts done by several persons in furtherance of common intention","m":0},
        "35":{"d":"acts done with criminal knowledge or intention","m":0},
        "107":{"d":"abetment of a thing","m":0},
        "108":{"d":"abettor","m":0},
        "109":{"d":"punishment of abetment if the act abetted is committed in consequence and where no express provision is made for its punishment","m":0},
        "110":{"d":"punishment of abetment if person abetted does act with different intention from that of abettor","m":0},
        "111":{"d":"liability of abettor when one act abetted and different act done","m":0},
        "114":{"d":"abettor present when offence is committed","m":0},
        "115":{"d":"abetment of offence punishable with death or imprisonment for life","m":84},
        "116":{"d":"abetment of offence punishable with imprisonment","m":0},
        "117":{"d":"abetting commission of offence by the public or by more than ten persons","m":36},
        "120A":{"d":"criminal conspiracy","m":0},
        "120B":{"d":"punishment of criminal conspiracy","m":9999},
        "143":{"d":"punishment for unlawful assembly","m":6},
        "144":{"d":"joining unlawful assembly armed with deadly weapon","m":24},
        "145":{"d":"joining or continuing in unlawful assembly knowing it has been commanded to disperse","m":24},
        "146":{"d":"rioting","m":0},
        "147":{"d":"punishment for rioting","m":24},
        "148":{"d":"rioting armed with deadly weapon","m":36},
        "149":{"d":"every member of unlawful assembly guilty of offence committed in prosecution of common objects","m":0},
        "150":{"d":"hiring or conniving at hiring of persons to join unlawful assembly","m":36},
        "151":{"d":"knowingly joining or continuing in assembly of five or more persons after it has been commanded to disperse","m":6},
        "152":{"d":"assaulting or obstructing public servant when suppressing riot","m":36},
        "153":{"d":"wantonly giving provocation with intent to cause riot","m":12},
        "153A":{"d":"promoting enmity between different groups","m":36},
        "153B":{"d":"imputations assertions prejudicial to national integration","m":36},
        "159":{"d":"affray","m":0},
        "160":{"d":"punishment for committing affray","m":1},
        "161":{"d":"public servant taking gratification other than legal remuneration","m":36},
        "166":{"d":"public servant disobeying law with intent to cause injury to any person","m":12},
        "166A":{"d":"public servant disobeying direction under law","m":24},
        "167":{"d":"public servant framing an incorrect document with intent to cause injury","m":36},
        "171B":{"d":"bribery","m":12},
        "171E":{"d":"punishment for bribery","m":12},
        "172":{"d":"absconding to avoid service of summons or other proceeding","m":1},
        "173":{"d":"preventing service of summons or other proceeding or preventing publication thereof","m":6},
        "174":{"d":"non-attendance in obedience to an order from public servant","m":1},
        "175":{"d":"omission to produce document to public servant by person legally bound to produce it","m":6},
        "176":{"d":"omission to give notice or information to public servant by person legally bound to give it","m":1},
        "177":{"d":"furnishing false information","m":6},
        "178":{"d":"refusing oath or affirmation when duly required by public servant to make it","m":6},
        "179":{"d":"refusing to answer public servant authorised to question","m":6},
        "180":{"d":"refusing to sign statement","m":3},
        "181":{"d":"false statement on oath or affirmation to public servant or person authorised to administer an oath or affirmation","m":36},
        "182":{"d":"false information with intent to cause public servant to use his lawful power to the injury of another person","m":6},
        "183":{"d":"resistance to the taking of property by the lawful authority of a public servant","m":6},
        "184":{"d":"obstructing sale of property offered for sale by authority of public servant","m":12},
        "185":{"d":"illegal purchase or bid for property offered for sale by authority of public servant","m":12},
        "186":{"d":"obstructing public servant in discharge of public functions","m":3},
        "187":{"d":"omission to assist public servant when bound by law to give assistance","m":12},
        "188":{"d":"disobedience to order duly promulgated by public servant","m":6},
        "189":{"d":"threat of injury to public servant","m":24},
        "190":{"d":"threat of injury to induce person to refrain from applying for protection to public servant","m":12},
        "191":{"d":"giving false evidence","m":0},
        "192":{"d":"fabricating false evidence","m":0},
        "193":{"d":"punishment for false evidence","m":84},
        "194":{"d":"giving or fabricating false evidence with intent to procure conviction of capital offence","m":9999},
        "195":{"d":"giving or fabricating false evidence with intent to procure conviction of offence punishable with imprisonment for life or imprisonment","m":84},
        "196":{"d":"using evidence known to be false","m":84},
        "197":{"d":"issuing or signing false certificate","m":84},
        "198":{"d":"using as true a certificate known to be false","m":84},
        "199":{"d":"false statement made in declaration which is by law receivable as evidence","m":84},
        "200":{"d":"using as true such declaration knowing it to be false","m":84},
        "201":{"d":"causing disappearance of evidence of offence or giving false information to screen offender","m":84},
        "202":{"d":"intentional omission to give information of offence by person bound to inform","m":6},
        "203":{"d":"giving false information respecting an offence committed","m":24},
        "204":{"d":"destruction of document to prevent its production as evidence","m":24},
        "205":{"d":"false personation for purpose of act or proceeding in suit or prosecution","m":24},
        "206":{"d":"fraudulent removal or concealment of property to prevent its seizure as a forfeiture or in execution","m":24},
        "207":{"d":"fraudulent claim to property to prevent its seizure as forfeiture or in execution","m":24},
        "208":{"d":"fraudulently suffering decree for sum not due","m":24},
        "209":{"d":"dishonestly making false claim in Court","m":24},
        "210":{"d":"fraudulently obtaining decree for sum not due","m":24},
        "211":{"d":"false charge of offence made with intent to injure","m":84},
        "212":{"d":"harbouring offender","m":60},
        "213":{"d":"taking gift etc. to screen an offender from punishment","m":84},
        "214":{"d":"offering gift or restoration of property in consideration of screening offender","m":84},
        "215":{"d":"taking gift to help to recover stolen property etc.","m":24},
        "216":{"d":"harbouring offender who has escaped from custody or whose apprehension has been ordered","m":36},
        "216A":{"d":"penalty for harbouring robbers or dacoits","m":84},
        "217":{"d":"public servant disobeying direction of law with intent to save person from punishment or property from forfeiture","m":24},
        "218":{"d":"public servant framing incorrect record or writing with intent to save person from punishment or property from forfeiture","m":36},
        "219":{"d":"public servant in judicial proceeding corruptly making report etc. contrary to law","m":84},
        "220":{"d":"commitment for trial or confinement by person having authority who knows that he is acting contrary to law","m":84},
        "221":{"d":"intentional omission to apprehend on the part of public servant bound to apprehend","m":84},
        "222":{"d":"intentional omission to apprehend on the part of public servant bound to apprehend person under sentence or lawfully committed","m":9999},
        "223":{"d":"escape from confinement or custody negligently suffered by public servant","m":36},
        "224":{"d":"resistance or obstruction by a person to his lawful apprehension","m":24},
        "225":{"d":"resistance or obstruction to lawful apprehension of another person","m":24},
        "225A":{"d":"omission to apprehend or sufferance of escape on the part of a public servant in cases not otherwise provided for","m":36},
        "225B":{"d":"resistance or obstruction to lawful apprehension or escape or rescue in cases not otherwise provided for","m":6},
        "228":{"d":"intentional insult or interruption to public servant sitting in judicial proceeding","m":6},
        "231":{"d":"counterfeiting coin","m":84},
        "236":{"d":"abetting in India the counterfeiting out of India of coin","m":36},
        "239":{"d":"delivery of coin possessed with knowledge that it is counterfeit","m":36},
        "240":{"d":"delivery of Indian coin possessed with knowledge that it is counterfeit","m":84},
        "243":{"d":"possession of counterfeit coin by person who knew it to be counterfeit when he became possessed thereof","m":60},
        "255":{"d":"counterfeiting government stamp","m":9999},
        "263A":{"d":"prohibition of fictitious stamps","m":3},
        "269":{"d":"negligently doing act known to be likely to spread infection of disease dangerous to life","m":6},
        "270":{"d":"malignantly doing act known to be likely to spread infection of disease dangerous to life","m":24},
        "272":{"d":"adulteration of food or drink intended for sale","m":6},
        "273":{"d":"sale of noxious food or drink","m":6},
        "274":{"d":"adulteration of drugs","m":6},
        "275":{"d":"sale of adulterated drugs","m":6},
        "276":{"d":"sale of drug as a different drug or preparation","m":6},
        "277":{"d":"fouling water of public spring or reservoir","m":36},
        "278":{"d":"making atmosphere noxious to health","m":6},
        "279":{"d":"rash driving or riding on a public way","m":6},
        "280":{"d":"rash navigation of vessel","m":6},
        "281":{"d":"exhibition of false light mark or buoy","m":84},
        "282":{"d":"conveying person by water for hire in unsafe or overloaded vessel","m":60},
        "284":{"d":"negligent conduct with respect to poisonous substance","m":6},
        "288":{"d":"negligent conduct with respect to pulling down or repairing buildings","m":6},
        "290":{"d":"punishment for public nuisance in cases not otherwise provided for","m":1},
        "291":{"d":"continuance of nuisance after injunction to discontinue","m":6},
        "292":{"d":"sale etc. of obscene books etc.","m":24},
        "293":{"d":"sale etc. of obscene objects to young person","m":36},
        "294":{"d":"obscene acts and songs","m":3},
        "295":{"d":"injuring or defiling place of worship with intent to insult the religion of any class","m":24},
        "295A":{"d":"deliberate and malicious acts intended to outrage religious feelings of any class by insulting its religion or religious beliefs","m":36},
        "296":{"d":"disturbing religious assembly","m":12},
        "297":{"d":"trespassing on burial places etc.","m":12},
        "298":{"d":"uttering words etc. with deliberate intent to wound the religious feelings of any person","m":12},
    },
    "BNS":{
        "103":{"d":"murder","m":9999},
        "104":{"d":"murder by life-convict","m":9999},
        "105":{"d":"culpable homicide not amounting to murder","m":120},
        "106":{"d":"causing death by negligence","m":60},
        "106(1)":{"d":"causing death by negligence","m":60},
        "106(2)":{"d":"causing death by negligence by rash or negligent driving — hit and run","m":120},
        "107":{"d":"causing death of quick unborn child by act amounting to culpable homicide","m":120},
        "108":{"d":"abetment of suicide or attempt to commit suicide","m":120},
        "109":{"d":"attempt to murder","m":120},
        "110":{"d":"attempt to commit culpable homicide","m":36},
        "111":{"d":"organised crime","m":9999},
        "113":{"d":"terrorist act","m":9999},
        "115":{"d":"voluntarily causing hurt","m":12},
        "115(1)":{"d":"voluntarily causing hurt","m":12},
        "115(2)":{"d":"voluntarily causing grievous hurt","m":84},
        "116":{"d":"voluntarily causing hurt by dangerous weapons or means","m":36},
        "116(1)":{"d":"voluntarily causing hurt by dangerous weapons or means","m":36},
        "116(2)":{"d":"voluntarily causing grievous hurt by dangerous weapons or means","m":9999},
        "117":{"d":"voluntarily causing grievous hurt","m":84},
        "117(1)":{"d":"voluntarily causing grievous hurt","m":84},
        "117(2)":{"d":"voluntarily causing grievous hurt resulting in death","m":120},
        "117(3)":{"d":"voluntarily causing grievous hurt to extort property","m":120},
        "118":{"d":"voluntarily causing grievous hurt by dangerous weapons or means","m":9999},
        "118(1)":{"d":"voluntarily causing grievous hurt by dangerous weapons or means","m":9999},
        "118(2)":{"d":"voluntarily causing grievous hurt by dangerous weapons causing death","m":9999},
        "119":{"d":"voluntarily causing grievous hurt to extort property or to constrain to an illegal act","m":120},
        "120":{"d":"voluntarily causing hurt to extort confession or to compel restoration of property","m":84},
        "121":{"d":"assault or criminal force to woman with intent to outrage her modesty","m":24},
        "121(1)":{"d":"assault or criminal force to woman with intent to outrage her modesty","m":24},
        "122":{"d":"sexual harassment","m":36},
        "123":{"d":"assault or use of criminal force to woman with intent to disrobe","m":84},
        "124":{"d":"voyeurism","m":84},
        "125":{"d":"stalking","m":60},
        "126":{"d":"wrongful confinement","m":12},
        "126(1)":{"d":"wrongful confinement","m":12},
        "126(2)":{"d":"wrongful confinement for three or more days","m":24},
        "126(3)":{"d":"wrongful confinement for ten or more days","m":36},
        "127":{"d":"wrongful restraint","m":1},
        "128":{"d":"kidnapping","m":84},
        "129":{"d":"abduction","m":0},
        "130":{"d":"kidnapping or abducting in order to murder","m":9999},
        "131":{"d":"kidnapping for ransom","m":9999},
        "132":{"d":"kidnapping or abducting with intent to confine person","m":84},
        "133":{"d":"kidnapping or abducting child under ten years of age","m":84},
        "137":{"d":"kidnapping or abducting woman to compel her marriage","m":120},
        "138":{"d":"trafficking of person","m":84},
        "64":{"d":"rape","m":9999},
        "64(1)":{"d":"rape","m":9999},
        "64(2)":{"d":"aggravated rape","m":9999},
        "65":{"d":"punishment for rape in certain cases — woman under 16 years","m":9999},
        "66":{"d":"rape causing death or vegetative state","m":9999},
        "67":{"d":"gang rape on woman under twelve years","m":9999},
        "70":{"d":"gang rape","m":9999},
        "70(1)":{"d":"gang rape","m":9999},
        "70(2)":{"d":"gang rape on woman under 18 years","m":9999},
        "71":{"d":"punishment for repeat offenders","m":9999},
        "74":{"d":"assault or criminal force to woman with intent to disrobe","m":84},
        "76":{"d":"voyeurism","m":84},
        "77":{"d":"stalking","m":60},
        "78":{"d":"word gesture or act intended to insult the modesty of a woman","m":36},
        "79":{"d":"sexual harassment","m":36},
        "80":{"d":"assault or criminal force to woman with intent to outrage her modesty","m":24},
        "85":{"d":"husband or relative of husband subjecting woman to cruelty","m":36},
        "86":{"d":"dowry death","m":9999},
        "87":{"d":"cohabitation caused by a man deceitfully inducing a belief of lawful marriage","m":120},
        "88":{"d":"marrying again during lifetime of husband or wife","m":84},
        "152":{"d":"act endangering sovereignty unity and integrity of India","m":9999},
        "191":{"d":"rioting","m":24},
        "191(1)":{"d":"unlawful assembly","m":6},
        "191(2)":{"d":"rioting","m":24},
        "192":{"d":"rioting armed with deadly weapon","m":36},
        "190":{"d":"every member of unlawful assembly guilty of offence committed in prosecution of common objects","m":0},
        "196":{"d":"promoting enmity between different groups","m":36},
        "197":{"d":"imputations assertions prejudicial to national integration","m":36},
        "281":{"d":"rash driving or riding on a public way","m":6},
        "303":{"d":"theft","m":36},
        "303(2)":{"d":"theft in a building tent or vessel","m":84},
        "304":{"d":"theft in dwelling house","m":84},
        "309":{"d":"robbery","m":120},
        "310":{"d":"dacoity","m":9999},
        "311":{"d":"robbery or dacoity with attempt to cause death or grievous hurt","m":9999},
        "314":{"d":"criminal breach of trust","m":36},
        "315":{"d":"criminal breach of trust by carrier merchant or agent","m":84},
        "316":{"d":"cheating","m":84},
        "316(2)":{"d":"cheating with knowledge that wrongful loss may ensue","m":36},
        "317":{"d":"cheating by personation","m":36},
        "318":{"d":"cheating and dishonestly inducing delivery of property or alteration in property","m":84},
        "319":{"d":"dishonest or fraudulent removal or concealment of property to prevent its seizure","m":24},
        "324":{"d":"mischief","m":36},
        "326":{"d":"mischief by fire or explosive substance","m":9999},
        "329":{"d":"forgery","m":24},
        "334":{"d":"forgery of valuable security or will","m":9999},
        "336":{"d":"forgery for purpose of cheating","m":84},
        "338":{"d":"forgery for purpose of harming reputation","m":36},
        "340":{"d":"using as genuine a forged document or electronic record","m":24},
        "351":{"d":"criminal intimidation","m":24},
        "351(1)":{"d":"criminal intimidation","m":24},
        "351(2)":{"d":"criminal intimidation — threat of death or grievous hurt or destruction of property","m":84},
        "351(3)":{"d":"criminal intimidation causing victim to do any act not legally bound to do","m":24},
        "352":{"d":"intentional insult with intent to provoke breach of the peace","m":24},
        "353":{"d":"statements conducing to public mischief","m":36},
        "356":{"d":"defamation","m":24},
        "3(5)":{"d":"acts done by several persons in furtherance of common intention","m":0},
        "3(5)(a)":{"d":"acts done by several persons in furtherance of common intention","m":0},
        "61":{"d":"criminal conspiracy","m":9999},
        "49":{"d":"abetment of a thing","m":0},
    },
    # Motor Vehicle Act 1988
    "MV ACT":{
        "134":{"d":"failure to stop vehicle, render aid and report accident","m":6},
        "134(a)":{"d":"failure to stop vehicle after accident and render aid to injured","m":6},
        "134(b)":{"d":"failure to report accident to police officer","m":6},
        "134(a)(b)":{"d":"failure to stop vehicle after accident, render aid to injured and report accident to police","m":6},
        "134(a)&(b)":{"d":"failure to stop vehicle after accident, render aid to injured and report accident to police","m":6},
        "184":{"d":"driving dangerously","m":12},
        "185":{"d":"driving by a drunken person or by a person under the influence of drugs","m":24},
        "186":{"d":"driving when mentally or physically unfit to drive","m":3},
        "187":{"d":"failure to comply with requirements of sections 132 134 and 186","m":3},
        "189":{"d":"racing and trials of speed","m":12},
        "192":{"d":"using vehicle without registration","m":12},
        "192A":{"d":"using vehicle without permit","m":12},
        "193":{"d":"agents and canvassers without proper authority","m":12},
        "194":{"d":"driving vehicle exceeding permissible weight","m":12},
        "196":{"d":"driving uninsured vehicle","m":3},
        "197":{"d":"taking vehicle without authority","m":3},
        "198":{"d":"interference with vehicle","m":3},
        "199":{"d":"offences relating to licensing of drivers","m":3},
        "201":{"d":"obstruction of traffic","m":1},
        "206":{"d":"failure to comply with direction of police officer","m":3},
    },
    # NDPS Act 1985
    "NDPS":{
        "8":{"d":"prohibition of certain operations","m":0},
        "15":{"d":"punishment for contravention in relation to poppy straw","m":120},
        "16":{"d":"punishment for contravention in relation to coca plant and coca leaves","m":84},
        "17":{"d":"punishment for contravention in relation to prepared opium","m":120},
        "18":{"d":"punishment for contravention in relation to opium poppy and opium","m":120},
        "19":{"d":"punishment for embezzlement of opium by cultivator","m":120},
        "20":{"d":"punishment for contravention in relation to cannabis plant and cannabis","m":120},
        "20(a)":{"d":"cultivation of cannabis plant","m":120},
        "20(b)":{"d":"production supply sale purchase transport import export or use of cannabis","m":120},
        "20(b)(i)":{"d":"small quantity of cannabis","m":6},
        "20(b)(ii)":{"d":"cannabis other than small quantity","m":120},
        "21":{"d":"punishment for contravention in relation to manufactured drugs and preparations","m":120},
        "21(a)":{"d":"small quantity of manufactured drugs","m":12},
        "21(b)":{"d":"manufactured drugs other than small quantity","m":120},
        "21(c)":{"d":"commercial quantity of manufactured drugs","m":120},
        "22":{"d":"punishment for contravention in relation to psychotropic substances","m":120},
        "23":{"d":"punishment for illegal import in India export from India or transhipment","m":120},
        "24":{"d":"punishment for external dealings","m":9999},
        "25":{"d":"punishment for allowing premises to be used for commission of an offence","m":120},
        "25A":{"d":"failure to maintain accounts or to submit information or returns","m":36},
        "26":{"d":"punishment for offences in relation to licences permits and authorisations","m":36},
        "27":{"d":"punishment for consumption of any narcotic drug or psychotropic substance","m":12},
        "27A":{"d":"punishment for financing illicit traffic and harbouring offenders","m":9999},
        "28":{"d":"punishment for attempts to commit offences","m":120},
        "29":{"d":"punishment for abetment and criminal conspiracy","m":120},
        "31":{"d":"enhanced punishment for offences after previous conviction","m":9999},
        "31A":{"d":"death penalty for certain offences after previous conviction","m":9999},
        "35":{"d":"presumption of culpable mental state","m":0},
        "37":{"d":"offences to be cognizable and non-bailable","m":0},
    },
    # Arms Act 1959
    "ARMS ACT":{
        "3":{"d":"acquisition and possession of firearms and ammunition","m":0},
        "5":{"d":"manufacture conversion repair test or proof of arms or ammunition","m":0},
        "7":{"d":"shortening of barrel of a firearm","m":0},
        "25":{"d":"punishment for certain offences relating to arms","m":36},
        "25(1)":{"d":"illegal acquisition possession manufacture sale transfer conversion repair test or proof of arms or ammunition","m":36},
        "25(1A)":{"d":"illegal possession of prohibited arms or prohibited ammunition","m":84},
        "25(1B)":{"d":"carrying firearms in public place","m":36},
        "25(2)":{"d":"fraudulently altering or tampering with identification marks on arms","m":36},
        "26":{"d":"punishment for contravention of conditions of licences","m":36},
        "27":{"d":"punishment for using arms etc. in contravention of section 5","m":84},
        "28":{"d":"punishment for using arms in certain cases","m":84},
        "29":{"d":"punishment for knowingly purchasing arms from unlicensed person","m":36},
        "30":{"d":"punishment for contravention of provisions of Act or rules","m":36},
    },
    # SC/ST (Prevention of Atrocities) Act 1989
    "SC/ST ACT":{
        "3":{"d":"punishment for offences of atrocities","m":60},
        "3(1)":{"d":"atrocity against member of Scheduled Caste or Scheduled Tribe","m":60},
        "3(1)(r)":{"d":"intentionally insults or intimidates with intent to humiliate a member of SC/ST in any place within public view","m":60},
        "3(1)(s)":{"d":"abuses any member of SC/ST by caste name in any place within public view","m":60},
        "3(1)(w)":{"d":"intentionally touches a woman belonging to SC/ST without her consent knowing that she belongs to SC/ST","m":60},
        "3(1)(x)":{"d":"uses words makes sounds or gestures or exhibits any object intending to insult modesty of woman belonging to SC/ST","m":60},
        "3(2)":{"d":"offences against SC/ST by person not a member of SC/ST","m":9999},
        "3(2)(v)":{"d":"commits offence punishable with imprisonment for a term of ten years or more against a person knowing that such person is a member of SC/ST","m":120},
        "3(2)(va)":{"d":"commits an offence specified in the Schedule against a person knowing that such person is a member of SC/ST","m":120},
        "4":{"d":"punishment for neglect of duties","m":12},
        "14":{"d":"Special Court","m":0},
        "18":{"d":"section 438 of the Code of Criminal Procedure not to apply to persons committing an offence under the Act","m":0},
    },
    # Dowry Prohibition Act 1961
    "DP ACT":{
        "3":{"d":"penalty for giving or taking dowry","m":60},
        "4":{"d":"penalty for demanding dowry","m":24},
        "4A":{"d":"ban on advertisement","m":6},
    },
    # Prevention of Corruption Act 1988
    "PC ACT":{
        "7":{"d":"public servant taking gratification other than legal remuneration in respect of an official act","m":84},
        "7A":{"d":"taking undue advantage to influence public servant","m":84},
        "8":{"d":"giving or taking undue advantage to or by persons concerned with cases before authority","m":84},
        "9":{"d":"taking undue advantage by public servant to influence superiors","m":84},
        "10":{"d":"abetment of offences defined in sections 7 and 11","m":60},
        "11":{"d":"public servant obtaining undue advantage without consideration from person concerned in proceeding or business transacted by such public servant","m":36},
        "12":{"d":"criminal misconduct by a public servant","m":0},
        "13":{"d":"criminal misconduct by a public servant","m":120},
        "13(1)":{"d":"criminal misconduct by a public servant","m":120},
        "14":{"d":"punishment for habitually committing offence under certain sections","m":9999},
        "15":{"d":"punishment for attempt","m":36},
    },
    # POCSO Act 2012
    "POCSO":{
        "3":{"d":"penetrative sexual assault","m":0},
        "4":{"d":"punishment for penetrative sexual assault","m":120},
        "4(1)":{"d":"penetrative sexual assault","m":120},
        "4(2)":{"d":"penetrative sexual assault on child below 16 years","m":9999},
        "5":{"d":"aggravated penetrative sexual assault","m":9999},
        "5(a)":{"d":"aggravated penetrative sexual assault by police officer","m":9999},
        "5(b)":{"d":"aggravated penetrative sexual assault by public servant","m":9999},
        "5(c)":{"d":"aggravated penetrative sexual assault by armed forces or security forces","m":9999},
        "5(d)":{"d":"aggravated penetrative sexual assault by management or staff of an institution","m":9999},
        "5(e)":{"d":"aggravated penetrative sexual assault on a child who is mentally ill","m":9999},
        "5(f)":{"d":"aggravated penetrative sexual assault causing grievous hurt or bodily harm or injury to the sexual organs of the child","m":9999},
        "5(g)":{"d":"aggravated penetrative sexual assault on a pregnant child","m":9999},
        "5(h)":{"d":"aggravated penetrative sexual assault on a child below the age of 12 years","m":9999},
        "5(i)":{"d":"aggravated penetrative sexual assault by a relative of the child","m":9999},
        "5(j)":{"d":"aggravated penetrative sexual assault causing the child to become pregnant","m":9999},
        "5(k)":{"d":"aggravated penetrative sexual assault on a child more than once or repeatedly","m":9999},
        "5(l)":{"d":"aggravated penetrative sexual assault on a child below sixteen years","m":9999},
        "5(m)":{"d":"aggravated penetrative sexual assault by person in a position of trust or authority","m":9999},
        "5(n)":{"d":"commission of penetrative sexual assault on a child by more than one person","m":9999},
        "6":{"d":"punishment for aggravated penetrative sexual assault","m":9999},
        "6(1)":{"d":"aggravated penetrative sexual assault","m":9999},
        "6(2)":{"d":"aggravated penetrative sexual assault on child below 16 years","m":9999},
        "7":{"d":"sexual assault","m":60},
        "8":{"d":"punishment for sexual assault","m":60},
        "9":{"d":"aggravated sexual assault","m":84},
        "9(a)":{"d":"aggravated sexual assault by police officer","m":84},
        "9(b)":{"d":"aggravated sexual assault by public servant","m":84},
        "9(l)":{"d":"aggravated sexual assault on child below 12 years","m":84},
        "10":{"d":"punishment for aggravated sexual assault","m":84},
        "11":{"d":"sexual harassment of the child","m":36},
        "12":{"d":"punishment for sexual harassment","m":36},
        "13":{"d":"use of child for pornographic purposes","m":60},
        "14":{"d":"punishment for use of child for pornographic purposes","m":60},
        "15":{"d":"punishment for storage of pornographic material involving child","m":36},
        "16":{"d":"abetment of an offence","m":0},
        "17":{"d":"punishment for abetment","m":0},
        "18":{"d":"punishment for attempt to commit an offence","m":0},
        "19":{"d":"reporting of offences","m":0},
    },
    # Prohibition of Child Marriage Act 2006
    "PCMA":{
        "9":{"d":"punishment for male adult marrying a child","m":24},
        "10":{"d":"punishment for solemnising a child marriage","m":24},
        "11":{"d":"punishment for promoting or permitting solemnisation of child marriages","m":24},
    },
    # Indian Evidence Act 1872 / Bharatiya Sakshya Adhiniyam 2023
    "BSA":{
        "23":{"d":"making false statement on oath — Bharatiya Sakshya Adhiniyam","m":84},
        "24":{"d":"fabricating false evidence — Bharatiya Sakshya Adhiniyam","m":84},
    },
    # Explosives Act 1884
    "EXPLOSIVES ACT":{
        "4":{"d":"making or possessing explosives in contravention","m":36},
        "5":{"d":"manufacture sale transport possession use etc. of explosives","m":120},
        "9B":{"d":"enhanced punishment for certain offences","m":9999},
    },
}

# Sections never framed as standalone charges — only appear as r/w
RW_ONLY = {
    "34","35","109","110","111","114","116","117","143","144","145",
    "149","150","151","152","153","190","3(5)","3(5)(a)","3(5)(b)",
    "49","61(1)","120A","190","191(1)","16","17","18","3"
}

ORDINALS = ["Firstly","Secondly","Thirdly","Fourthly","Fifthly",
            "Sixthly","Seventhly","Eighthly","Ninthly","Tenthly"]

# ═══════════════════════════════════════════════════════════════

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
    if ftype == 'typed_pdf':
        if progress_cb: progress_cb("Reading PDF...")
        return extract_pdf_text(path)
    elif ftype == 'scanned_pdf':
        if progress_cb: progress_cb("Reading scanned PDF...")
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
EXTRACT_PROMPT = """You are a senior legal analyst for Indian criminal courts. You have complete knowledge of IPC, CrPC, BNS, BNSS, POCSO Act, NDPS Act, Arms Act, SC/ST Act, Motor Vehicle Act, Dowry Prohibition Act, Prevention of Corruption Act and all other Indian criminal laws.

READ ONLY the charge sheet portion. IGNORE completely: FIR copy, List of Witnesses / Memo of Evidence, Inquest Report, Post Mortem Report, MVI Report, Arrest Memo, any document after "Hence the charge" line or after the police officer signature.

The charge sheet starts with "IN THE COURT OF..." and ends with "Hence the charge." or "Thus the accused...committed an offence...Hence the charge."

Return ONLY a valid JSON object. No explanation. No markdown. Just JSON.

{
  "ps": "Police Station name from the line Sub-Inspector of Police [PS Name] P.S. or Sub-Divisional Police Officer [Place]",
  "accused": [
    {
      "name": "Full name exactly as written — include all initials and alias with @ symbol",
      "age": "Age as number string only — empty if not mentioned",
      "father": "Complete name after S/o or D/o or W/o — include all initials exactly — keep the word Late or late if it appears before the name in the charge sheet",
      "village": "VILLAGE RULES below",
      "mandal": "MANDAL RULES below",
      "district": "Word(s) immediately before District keyword — empty if not found",
      "full_addr": "Complete address of this accused from accused block",
      "exempted": false
    }
  ],
  "sec_line": "Exact section text from Charge sheet filed line — starting with U/Sec. or U/s. exactly as written including all sections, r/w clauses, and Act names. Example: U/s 376(1) IPC and Sec.6 r/w 5(l) of POCSO Act 2012",
  "law": "IPC or BNS — primary substantive law used",
  "sections": ["SECTION EXTRACTION RULES below"],
  "acts": ["IPC", "POCSO"],
  "cr_no": "Case number like 45/2023",
  "incident_date": "Date of incident in DD.MM.YYYY format — or a period like August 2022 if no specific date — empty if not found",
  "incident_time": "Time of incident like 10.00 PM or 4.30 p.m. — empty if not found",
  "incident_place": "Complete specific place of occurrence as described in the charge sheet",
  "victim_name": "Name of deceased or victim if mentioned",
  "accused_act": "For EACH section separately — what the accused did relating to THAT specific section — used in framing of charges. One sentence per section.",
  "doctor_opinion": "Medical or forensic expert opinion if mentioned — PM findings cause of death potency test DNA report medical certificate — one or two sentences",
  "facts_brief": "FACTS BRIEF RULES below",
  "charge_descriptions": [
    {
      "sec": "exact section number as it appears in sections array",
      "act": "exact Act name like IPC or BNS or POCSO Act 2012 or Motor Vehicle Act 1988 or NDPS Act 1985 or Arms Act 1959 etc.",
      "one_line": "one specific sentence describing what the accused did for THIS section only — use formal court language — start with a verb — do NOT repeat same sentence for different sections"
    }
  ],
  "case_type": "SUMMONS or WARRANT — based on maximum punishment. More than 3 years = WARRANT. Sections 34 IPC 149 IPC 3(5) BNS on their own do not make it WARRANT.",
  "is_sessions": "true if ANY section is triable by Sessions Court or Special Court — murder rape dacoity POCSO SC/ST etc. — else false"
}

SECTION EXTRACTION RULES — CRITICAL:
1. Extract section numbers from the Charge sheet filed line ONLY
2. Split sections ONLY at commas and the word 'and' between completely separate sections
3. NEVER split at 'r/w' within a section — 'Sec.6 r/w 5(l)' is ONE section — extract as '6 r/w 5(l)'
4. The final r/w that applies to ALL sections (like 'r/w 34 IPC' or 'r/w 3(5) BNS' at the very end) is NOT a separate section — it is the global r/w clause — do NOT include it in the sections array as a standalone entry
5. Examples:
   - 'U/s 376(1) IPC and Sec.6 r/w 5(l) of POCSO Act 2012' → sections: ['376(1)', '6 r/w 5(l)']
   - 'U/s 118(1), 115(2), 351(2) r/w 3(5) BNS' → sections: ['118(1)', '115(2)', '351(2)'] and global rw is '3(5) BNS'
   - 'U/s 498A, 323 IPC r/w 34 IPC' → sections: ['498A', '323'] and global rw is '34 IPC'
   - 'U/s 20(b)(ii) of NDPS Act' → sections: ['20(b)(ii)']
   - 'U/s 106(1) BNS and Sec.134(a)(b) of MV Act 1988' → sections: ['106(1)', '134(a)(b)']

VILLAGE RULES (apply in order for each accused):
1. Word(s) immediately before Village keyword → use as village
2. Else word(s) before Town keyword → use as village
3. Else word(s) before Panchayat or Panchayath or Grama Panchayat → use as village
4. Else word(s) before H/w or H/wada or Hamlet → use as village including the H/w or H/wada suffix
5. Else if none above found → use same value as mandal for village

MANDAL RULES:
1. Word(s) immediately before Mandal keyword → use as mandal
2. Else word(s) before Taluqa or Taluka or Taluk → use as mandal
3. Else word(s) before Town and Municipality → use entire Town and Municipality phrase
4. Else empty string

FACTS BRIEF RULES:
Write a complete factual narrative in formal court English covering ALL of the following that are present in the charge sheet:
- Date time and place of occurrence
- Identity of accused and victim and their relationship if any
- Modus operandi — exactly how the accused committed the act — specific details
- What happened step by step
- Injuries sustained by victim
- If death case — how and where deceased died
- Medical treatment given — hospitals visited
- Post Mortem findings and cause of death if applicable
- Any forensic findings — DNA report potency test FSL report etc. if applicable
- Any special act violation — fleeing scene without aid reporting etc.
- Role of each accused if multiple accused

LENGTH RULE:
- Magistrate Court cases (normal IPC/BNS): Write a concise but complete narrative — cover all important points without unnecessary details — typically 5 to 8 sentences
- Sessions Court and Special Court cases (murder rape POCSO SC/ST dacoity etc.): Write a more elaborate narrative — cover every important point in detail including forensic and medical evidence — typically 8 to 15 sentences
- In both cases — every important fact must be covered — do NOT leave out doctor opinion PM findings forensic evidence

IMPORTANT:
- List EVERY accused — count before answering — if 3 accused the array must have 3 entries
- exempted: true if accused described as not charged exonerated no case made out
- charge_descriptions must have ONE entry per section in sections array — each one_line must be DIFFERENT and specific to that section
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

    # Ensure required fields
    if "accused" not in data or not data["accused"]:
        data["accused"] = [{"name":"","age":"","father":"","village":"",
            "mandal":"","district":"","full_addr":"","exempted":False}]
    for acc in data["accused"]:
        if "exempted" not in acc: acc["exempted"] = False
    defaults = {
        "sections":[],"acts":[],"law":"IPC","sec_line":"","ps":"",
        "cr_no":"","incident_date":"","incident_time":"","incident_place":"",
        "victim_name":"","accused_act":"","doctor_opinion":"",
        "facts_brief":"","charge_descriptions":[],"case_type":"","is_sessions":False
    }
    for k,v in defaults.items():
        if k not in data: data[k] = v
    return data

# ═══════════════════════════════════════════════════════════════
# CASE CLASSIFICATION
# ═══════════════════════════════════════════════════════════════
def lookup_section(sec, law):
    """Look up section in all databases — handles r/w sections."""
    # For r/w sections like "6 r/w 5(l)" — look up the main section "6"
    main_sec = sec.split(' r/w')[0].strip() if ' r/w' in sec else sec
    base = re.sub(r'\([^)]*\)','',main_sec).strip()
    for db_key in [law,"IPC","BNS","POCSO","MV ACT","NDPS","ARMS ACT","SC/ST ACT","DP ACT","PC ACT","EXPLOSIVES ACT","PCMA","BSA"]:
        db = DB.get(db_key,{})
        for v in [sec, main_sec, base, main_sec+"A", main_sec.replace("(","").replace(")","")]:
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
    return {"case_type":"WARRANT" if max_m>36 else "SUMMONS",
            "max_m":max_m,"found":found}

def get_exam_ref(law, ct, court_type):
    if court_type == "sessions":
        return "228 Cr.P.C." if law=="IPC" else "240 BNSS"
    if law == "BNS":
        return "262 BNSS" if ct=="SUMMONS" else "274 BNSS"
    return "251 Cr.P.C." if ct=="SUMMONS" else "239 Cr.P.C."

def extract_global_rw(sec_line):
    """Extract the global r/w clause that applies to ALL sections."""
    if not sec_line: return ""
    # Match r/w at end of sec_line (global r/w)
    m = re.search(r'\br/w\.?\s+([\d\(\)A-Za-z,\s&\.]+?(?:IPC|BNS|Act[\-\d\s]*))\s*$',sec_line.strip(),re.I)
    return re.sub(r'\s+',' ',m.group(1)).strip() if m else ""

def filter_substantive(sections):
    return [s for s in sections if s.strip() not in RW_ONLY and not s.strip().startswith('r/w')]

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
    p.paragraph_format.space_before=Pt(0); p.paragraph_format.space_after=Pt(0)
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

    doc  = new_doc()
    TOTAL=9000; DW=[2200,2600,1750,2450]; QW0=1050; QW1=TOTAL-QW0
    acc_list = [a for a in info.get("accused",[{}]) if not a.get("exempted",False)]
    if not acc_list: acc_list = info.get("accused",[{}])

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
    global_rw = extract_global_rw(sec_line)
    analysis= classify_case(raw_secs,law)
    idate   = info.get("incident_date","")
    itime   = info.get("incident_time","")
    iplace  = info.get("incident_place","")
    victim  = info.get("victim_name","")
    judge_str = judge if judge else f"Sri ___________________________, {d1} {d2}"

    # Build date-time-place string
    if idate and itime and iplace:
        dtp = f"on {idate} at about {itime}, {iplace}"
    elif idate and iplace:
        dtp = f"on {idate}, {iplace}"
    elif iplace:
        dtp = f"at {iplace}"
    else:
        dtp = "on the date, time and place mentioned in the charge sheet"

    # Per-section charge description lookup
    cd_map = {}
    for cd in info.get("charge_descriptions",[]):
        cd_map[cd.get("sec","")] = cd

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

    total=len(secs)
    for idx,sec in enumerate(secs):
        is_first=(idx==0); is_last=(idx==total-1)
        ordinal="Lastly" if is_last and total>1 else (ORDINALS[idx] if idx<len(ORDINALS) else f"{idx+1}thly")
        cognizance="within my cognizance" if is_first else "within the cognizance of this Court"

        cd_entry=cd_map.get(sec,{})
        act_name=cd_entry.get("act","") or law
        one_line=cd_entry.get("one_line","")

        # Build section reference
        if global_rw and not ' r/w ' in sec:
            sec_ref=f"Sec.{sec} r/w {global_rw}"
        elif ' r/w ' in sec:
            parts=sec.split(' r/w ',1)
            sec_ref=f"Sec.{parts[0].strip()} r/w {parts[1].strip()} of {act_name}" if act_name and act_name not in parts[1] else f"Sec.{sec}"
        else:
            sec_ref=f"Sec.{sec} of {act_name}"

        # Use per-section one_line if available
        if one_line:
            charge_body=one_line
        elif victim:
            charge_body=f"committed the offence against {victim}"
        else:
            entry,_=lookup_section(sec,law)
            charge_body=f"committed the offence of {entry['d']}" if entry else "committed the offence as stated in the charge sheet"

        charge_text=f"That you {dtp}, {charge_body} and that you thereby committed an offence punishable under {sec_ref} and {cognizance}."

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
    add_desig_block(doc,d1,d2)

    buf=io.BytesIO(); doc.save(buf); buf.seek(0)
    return buf.read()

# ═══════════════════════════════════════════════════════════════
# COMBINE DOCUMENTS
# ═══════════════════════════════════════════════════════════════
def combine_docs(exam_bytes, charges_bytes):
    exam_doc = DocxDoc(io.BytesIO(exam_bytes))
    charges_doc = DocxDoc(io.BytesIO(charges_bytes))
    last_para = exam_doc.paragraphs[-1] if exam_doc.paragraphs else exam_doc.add_paragraph()
    run = last_para.add_run()
    br = OxmlElement('w:br')
    br.set(qn('w:type'),'page')
    run._r.append(br)
    for element in charges_doc.element.body:
        exam_doc.element.body.append(copy.deepcopy(element))
    buf = io.BytesIO()
    exam_doc.save(buf); buf.seek(0)
    return buf.read()

# ═══════════════════════════════════════════════════════════════
# MAIN PROCESSING PIPELINE
# ═══════════════════════════════════════════════════════════════
def process_file(path, cfg, court_type, manual_ref, progress_cb=None):
    api_key = cfg.get("api_key","")
    if not api_key: raise Exception("API key not set. Click Settings.")

    orig_name = Path(path).stem
    ftype     = get_file_type(path)
    info      = ai_extract(path, ftype, api_key, progress_cb)

    raw_secs  = info.get("sections",[])
    law       = info.get("law","IPC")
    db_anal   = classify_case(raw_secs, law)
    db_ct     = db_anal["case_type"]
    ai_ct     = str(info.get("case_type","")).upper().strip()
    if ai_ct not in ["SUMMONS","WARRANT"]: ai_ct = db_ct
    ai_sessions = str(info.get("is_sessions","")).lower() == "true"
    ct = "WARRANT" if (ai_ct=="WARRANT" or db_ct=="WARRANT") else "SUMMONS"

    if manual_ref:
        exam_ref = manual_ref
        ct = "SUMMONS" if any(x in manual_ref for x in ["251","262"]) else "WARRANT"
    elif court_type == "sessions" or ai_sessions:
        court_type = "sessions"
        exam_ref = "228 Cr.P.C." if law=="IPC" else "240 BNSS"
        ct = "WARRANT"
    else:
        exam_ref = get_exam_ref(law, ct, court_type)

    info["case_type_resolved"] = ct
    all_acc    = info.get("accused",[{}])
    active_acc = [a for a in all_acc if not a.get("exempted",False)]
    if not active_acc: active_acc = all_acc
    info["accused"] = active_acc

    if progress_cb: progress_cb("Generating Word document...")

    exam_bytes   = make_examination_doc(info, cfg, court_type, exam_ref, ct)
    need_charges = (ct=="WARRANT" or court_type=="sessions") and filter_substantive(raw_secs)

    if need_charges:
        charges_bytes = make_charges_doc(info, cfg, court_type, law)
        combined      = combine_docs(exam_bytes, charges_bytes)
        docs = [{"fname": orig_name+".docx", "data": combined}]
    else:
        docs = [{"fname": orig_name+".docx", "data": exam_bytes}]

    return {
        "name":Path(path).name,"status":"ok","law":law,
        "sec_line":info.get("sec_line",""),"sections":raw_secs,"ct":ct,
        "exam_ref":exam_ref,"ps":info.get("ps",""),"cr_no":info.get("cr_no",""),
        "acc_count":len(info.get("accused",[])),"file_type":ftype,"docs":docs,
    }

# ═══════════════════════════════════════════════════════════════
# COURT TYPE + MODE SELECTION DIALOG
# ═══════════════════════════════════════════════════════════════
