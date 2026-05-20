from flask import Flask, render_template, request, send_file
import os
import re
import PyPDF2
import docx
from datetime import datetime
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                 TableStyle, HRFlowable, KeepTogether)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ============================================================
# 🗂️ ALL SKILLS — Domain wise
# ============================================================
ALL_SKILLS = {
    "html": "💻 IT", "css": "💻 IT", "javascript": "💻 IT",
    "react": "💻 IT", "angular": "💻 IT", "vue": "💻 IT",
    "node.js": "💻 IT", "nodejs": "💻 IT", "express.js": "💻 IT", "express": "💻 IT",
    "python": "💻 IT", "java": "💻 IT", "c++": "💻 IT", "c#": "💻 IT",
    "php": "💻 IT", "typescript": "💻 IT", "flask": "💻 IT", "django": "💻 IT",
    "spring boot": "💻 IT", "mongodb": "💻 IT", "mysql": "💻 IT",
    "postgresql": "💻 IT", "sqlite": "💻 IT", "sql": "💻 IT",
    "git": "💻 IT", "github": "💻 IT", "docker": "💻 IT", "kubernetes": "💻 IT",
    "aws": "💻 IT", "azure": "💻 IT", "linux": "💻 IT", "bootstrap": "💻 IT",
    "rest api": "💻 IT", "graphql": "💻 IT", "redux": "💻 IT",
    "next.js": "💻 IT", "tailwind": "💻 IT", "tailwind css": "💻 IT",
    "firebase": "💻 IT", "kotlin": "💻 IT", "android": "💻 IT",
    "selenium": "💻 IT", "postman": "💻 IT", "oops": "💻 IT",
    "data structures": "💻 IT", "algorithms": "💻 IT",
    "operating system": "💻 IT", "computer networks": "💻 IT",
    "dbms": "💻 IT", "sass": "💻 IT", "jquery": "💻 IT",
    "wordpress": "💻 IT", "asp.net": "💻 IT", "programming": "💻 IT",
    "software development": "💻 IT", "web development": "💻 IT",
    "app development": "💻 IT", "coding": "💻 IT",
    "machine learning": "📊 Data Science", "deep learning": "📊 Data Science",
    "data analysis": "📊 Data Science", "data science": "📊 Data Science",
    "pandas": "📊 Data Science", "numpy": "📊 Data Science",
    "tensorflow": "📊 Data Science", "keras": "📊 Data Science",
    "pytorch": "📊 Data Science", "scikit-learn": "📊 Data Science",
    "sklearn": "📊 Data Science",
    "natural language processing": "📊 Data Science",
    "computer vision": "📊 Data Science",
    "power bi": "📊 Data Science", "tableau": "📊 Data Science",
    "r programming": "📊 Data Science", "big data": "📊 Data Science",
    "hadoop": "📊 Data Science", "apache spark": "📊 Data Science",
    "data mining": "📊 Data Science", "data visualization": "📊 Data Science",
    "opencv": "📊 Data Science", "matplotlib": "📊 Data Science",
    "seaborn": "📊 Data Science", "statistics": "📊 Data Science",
    "artificial intelligence": "📊 Data Science",
    "anatomy": "🏥 Medical", "physiology": "🏥 Medical",
    "pharmacology": "🏥 Medical", "patient care": "🏥 Medical",
    "clinical skills": "🏥 Medical", "medical diagnosis": "🏥 Medical",
    "surgery": "🏥 Medical", "nursing": "🏥 Medical",
    "first aid": "🏥 Medical", "icu care": "🏥 Medical",
    "ecg": "🏥 Medical", "radiology": "🏥 Medical",
    "pathology": "🏥 Medical", "medical coding": "🏥 Medical",
    "health informatics": "🏥 Medical", "mbbs": "🏥 Medical",
    "bds": "🏥 Medical", "ayurveda": "🏥 Medical",
    "homeopathy": "🏥 Medical", "biochemistry": "🏥 Medical",
    "pediatrics": "🏥 Medical", "gynecology": "🏥 Medical",
    "dentistry": "🏥 Medical", "pharmacy": "🏥 Medical",
    "clinical research": "🏥 Medical", "healthcare": "🏥 Medical",
    "legal research": "⚖️ Law", "contract drafting": "⚖️ Law",
    "litigation": "⚖️ Law", "corporate law": "⚖️ Law",
    "criminal law": "⚖️ Law", "family law": "⚖️ Law",
    "intellectual property": "⚖️ Law", "taxation law": "⚖️ Law",
    "cyber law": "⚖️ Law", "constitutional law": "⚖️ Law",
    "legal writing": "⚖️ Law", "court proceedings": "⚖️ Law",
    "arbitration": "⚖️ Law", "mediation": "⚖️ Law",
    "llb": "⚖️ Law", "ba.llb": "⚖️ Law", "ballb": "⚖️ Law",
    "moot court": "⚖️ Law", "case analysis": "⚖️ Law",
    "case study": "⚖️ Law", "legal drafting": "⚖️ Law",
    "client communication": "⚖️ Law", "law intern": "⚖️ Law",
    "law internship": "⚖️ Law", "legal intern": "⚖️ Law",
    "legal internship": "⚖️ Law", "public speaking": "⚖️ Law",
    "ethical leadership": "⚖️ Law", "problem resolution": "⚖️ Law",
    "written communication": "⚖️ Law", "legal compliance": "⚖️ Law",
    "advocate": "⚖️ Law", "paralegal": "⚖️ Law", "legal aid": "⚖️ Law",
    "bar council": "⚖️ Law", "juris": "⚖️ Law", "law": "⚖️ Law",
    "accounting": "📈 Finance", "financial analysis": "📈 Finance",
    "tally": "📈 Finance", "tally erp": "📈 Finance",
    "gst": "📈 Finance", "income tax": "📈 Finance",
    "auditing": "📈 Finance", "budgeting": "📈 Finance",
    "financial reporting": "📈 Finance", "balance sheet": "📈 Finance",
    "investment banking": "📈 Finance", "stock market": "📈 Finance",
    "mutual funds": "📈 Finance", "risk management": "📈 Finance",
    "financial modeling": "📈 Finance", "cfa": "📈 Finance",
    "cpa": "📈 Finance", "quickbooks": "📈 Finance",
    "taxation": "📈 Finance", "bookkeeping": "📈 Finance",
    "finance": "📈 Finance", "banking": "📈 Finance",
    "photoshop": "🎨 Design", "adobe photoshop": "🎨 Design",
    "illustrator": "🎨 Design", "figma": "🎨 Design",
    "adobe xd": "🎨 Design", "canva": "🎨 Design",
    "ui/ux": "🎨 Design", "ui ux": "🎨 Design",
    "graphic design": "🎨 Design", "video editing": "🎨 Design",
    "premiere pro": "🎨 Design", "after effects": "🎨 Design",
    "3d modeling": "🎨 Design", "blender": "🎨 Design",
    "motion graphics": "🎨 Design", "brand design": "🎨 Design",
    "indesign": "🎨 Design", "corel draw": "🎨 Design",
    "wireframing": "🎨 Design", "prototyping": "🎨 Design",
    "autocad": "🏗️ Engineering", "solidworks": "🏗️ Engineering",
    "matlab": "🏗️ Engineering", "ansys": "🏗️ Engineering",
    "staad pro": "🏗️ Engineering", "revit": "🏗️ Engineering",
    "structural analysis": "🏗️ Engineering",
    "construction management": "🏗️ Engineering",
    "mechanical design": "🏗️ Engineering",
    "thermodynamics": "🏗️ Engineering",
    "fluid mechanics": "🏗️ Engineering",
    "surveying": "🏗️ Engineering", "bim": "🏗️ Engineering",
    "catia": "🏗️ Engineering", "vlsi": "🏗️ Engineering",
    "plc": "🏗️ Engineering", "scada": "🏗️ Engineering",
    "lesson planning": "📚 Education",
    "curriculum development": "📚 Education",
    "classroom management": "📚 Education",
    "student assessment": "📚 Education",
    "e-learning": "📚 Education", "educational technology": "📚 Education",
    "special education": "📚 Education",
    "b.ed": "📚 Education", "ctet": "📚 Education",
    "teaching": "📚 Education", "mentoring": "📚 Education",
    "tutoring": "📚 Education", "coaching": "📚 Education",
    "digital marketing": "📣 Marketing/HR", "seo": "📣 Marketing/HR",
    "social media marketing": "📣 Marketing/HR",
    "content writing": "📣 Marketing/HR", "email marketing": "📣 Marketing/HR",
    "google ads": "📣 Marketing/HR", "facebook ads": "📣 Marketing/HR",
    "brand management": "📣 Marketing/HR", "market research": "📣 Marketing/HR",
    "hr management": "📣 Marketing/HR", "recruitment": "📣 Marketing/HR",
    "payroll": "📣 Marketing/HR", "performance management": "📣 Marketing/HR",
    "employee relations": "📣 Marketing/HR", "copywriting": "📣 Marketing/HR",
    "bioinformatics": "🔬 Science", "scientific writing": "🔬 Science",
    "hypothesis testing": "🔬 Science", "biotechnology": "🔬 Science",
    "microbiology": "🔬 Science", "genetics": "🔬 Science",
    "laboratory": "🔬 Science", "research methodology": "🔬 Science",
    "ms office": "🌐 General", "microsoft office": "🌐 General",
    "microsoft excel": "🌐 General", "microsoft word": "🌐 General",
    "powerpoint": "🌐 General", "communication skills": "🌐 General",
    "leadership": "🌐 General", "teamwork": "🌐 General",
    "project management": "🌐 General", "time management": "🌐 General",
    "problem solving": "🌐 General", "problem-solving": "🌐 General",
    "analytical thinking": "🌐 General", "critical thinking": "🌐 General",
    "computer skills": "🌐 General",
}

# ============================================================
# 🗂️ JOB DATABASE
# ============================================================
JOB_DATABASE = [
    {"job": "Frontend Developer",        "required": ["html","css","javascript","react","bootstrap"]},
    {"job": "Backend Developer",         "required": ["python","flask","django","sql","rest api"]},
    {"job": "Full Stack Developer",      "required": ["html","css","javascript","react","nodejs","mongodb","python"]},
    {"job": "Java Developer",            "required": ["java","spring boot","sql","git","rest api"]},
    {"job": "DevOps Engineer",           "required": ["docker","kubernetes","linux","aws","git"]},
    {"job": "Cloud Engineer",            "required": ["aws","azure","docker","linux","python"]},
    {"job": "Android Developer",         "required": ["java","android","kotlin","sql","git"]},
    {"job": "Software Tester / QA",      "required": ["selenium","python","sql","git"]},
    {"job": "UI/UX Designer",            "required": ["figma","adobe xd","wireframing","prototyping","canva"]},
    {"job": "Data Analyst",              "required": ["python","pandas","numpy","sql","power bi"]},
    {"job": "Data Scientist",            "required": ["machine learning","python","pandas","tensorflow"]},
    {"job": "ML Engineer",               "required": ["machine learning","deep learning","tensorflow","pytorch","python"]},
    {"job": "Business Intelligence",     "required": ["power bi","tableau","sql","data analysis","microsoft excel"]},
    {"job": "Medical Doctor (MBBS)",     "required": ["anatomy","physiology","pharmacology","medical diagnosis","patient care"]},
    {"job": "Nurse",                     "required": ["patient care","nursing","first aid","anatomy","physiology"]},
    {"job": "Dentist (BDS)",             "required": ["bds","anatomy","patient care","surgery","dentistry"]},
    {"job": "Pharmacist",                "required": ["pharmacology","biochemistry","patient care"]},
    {"job": "Medical Coder",             "required": ["medical coding","health informatics","anatomy"]},
    {"job": "Corporate Lawyer",          "required": ["corporate law","contract drafting","legal research","legal writing"]},
    {"job": "Criminal Lawyer",           "required": ["criminal law","litigation","court proceedings","legal research"]},
    {"job": "Law Intern",                "required": ["llb","case analysis","legal research","contract drafting","client communication"]},
    {"job": "Legal Advisor",             "required": ["legal research","arbitration","mediation","legal writing"]},
    {"job": "Tax Consultant",            "required": ["taxation law","gst","income tax","legal writing"]},
    {"job": "Paralegal",                 "required": ["legal research","legal drafting","case analysis","written communication"]},
    {"job": "Chartered Accountant",      "required": ["accounting","auditing","gst","income tax","tally"]},
    {"job": "Financial Analyst",         "required": ["financial analysis","microsoft excel","financial modeling","investment banking"]},
    {"job": "Accountant",                "required": ["accounting","tally","gst","microsoft excel","balance sheet"]},
    {"job": "Graphic Designer",          "required": ["photoshop","illustrator","canva","graphic design"]},
    {"job": "Video Editor",              "required": ["premiere pro","after effects","video editing","motion graphics"]},
    {"job": "Civil Engineer",            "required": ["autocad","staad pro","structural analysis","surveying"]},
    {"job": "Mechanical Engineer",       "required": ["solidworks","catia","ansys","matlab","mechanical design"]},
    {"job": "School Teacher",            "required": ["teaching","lesson planning","classroom management","student assessment"]},
    {"job": "Corporate Trainer",         "required": ["communication skills","leadership","microsoft office"]},
    {"job": "Digital Marketing Manager", "required": ["digital marketing","seo","social media marketing","google ads","content writing"]},
    {"job": "HR Manager",                "required": ["hr management","recruitment","payroll","employee relations"]},
    {"job": "Content Writer",            "required": ["content writing","seo","communication skills"]},
    {"job": "Research Scientist",        "required": ["scientific writing","hypothesis testing","bioinformatics"]},
    {"job": "Biotechnologist",           "required": ["biotechnology","genetics","microbiology"]},
    {"job": "AI Engineer",               "required": ["artificial intelligence","deep learning","python","tensorflow","natural language processing"]},
    {"job": "Data Engineer",             "required": ["python","sql","big data","apache spark","hadoop"]},
    {"job": "NLP Engineer",              "required": ["natural language processing","python","deep learning","tensorflow","scikit-learn"]},
    {"job": "Computer Vision Engineer",  "required": ["computer vision","opencv","deep learning","pytorch","python"]},
    {"job": "Statistician",              "required": ["statistics","r programming","data analysis","python","data visualization"]},
    {"job": "IP Lawyer",                 "required": ["intellectual property","legal research","contract drafting","litigation","legal writing"]},
    {"job": "Cyber Law Specialist",      "required": ["cyber law","legal research","legal writing","constitutional law"]},
    {"job": "Family Lawyer",             "required": ["family law","litigation","court proceedings","mediation","legal research"]},
    {"job": "Tax Analyst",               "required": ["income tax","gst","taxation","tally","accounting"]},
    {"job": "Audit Associate",           "required": ["auditing","accounting","financial reporting","tally","ms office"]},
    {"job": "Stock Market Analyst",      "required": ["stock market","investment banking","financial analysis","mutual funds","risk management"]},
    {"job": "Brand Designer",            "required": ["brand design","illustrator","photoshop","canva","graphic design"]},
    {"job": "Motion Designer",           "required": ["motion graphics","after effects","premiere pro","blender"]},
    {"job": "Product Designer",          "required": ["figma","wireframing","prototyping","ui/ux","adobe xd"]},
    {"job": "Structural Engineer",       "required": ["structural analysis","autocad","staad pro","revit","surveying"]},
    {"job": "VLSI Engineer",             "required": ["vlsi","matlab","autocad","electrical engineering"]},
    {"job": "Online Tutor",              "required": ["tutoring","e-learning","educational technology","communication skills"]},
    {"job": "Curriculum Designer",       "required": ["curriculum development","lesson planning","educational technology","student assessment"]},
    {"job": "Social Media Manager",      "required": ["social media marketing","content writing","digital marketing","copywriting","brand management"]},
    {"job": "Recruitment Specialist",    "required": ["recruitment","hr management","employee relations","communication skills"]},
    {"job": "HR Business Partner",       "required": ["hr management","performance management","employee relations","payroll","recruitment"]},
    {"job": "Microbiologist",            "required": ["microbiology","laboratory","genetics","biochemistry","scientific writing"]},
    {"job": "Lab Technician",            "required": ["laboratory","microbiology","biochemistry","research methodology"]},
    {"job": "Clinical Research Assoc.",  "required": ["clinical research","scientific writing","healthcare","research methodology"]},
    {"job": "Physiotherapist",           "required": ["anatomy","physiology","patient care","clinical skills","first aid"]},
    {"job": "Bioinformatics Analyst",    "required": ["bioinformatics","python","data analysis","genetics"]},
]

DOMAIN_MAP = {
    "💻 IT / Web Development":       "💻 IT",
    "📊 Data Science / ML / AI":     "📊 Data Science",
    "🏥 Medical / Healthcare":        "🏥 Medical",
    "⚖️ Law / Legal":                "⚖️ Law",
    "📈 Finance / Accounting":        "📈 Finance",
    "🎨 Design / Creative":           "🎨 Design",
    "🏗️ Civil / Mech. Engineering":  "🏗️ Engineering",
    "📚 Education / Teaching":        "📚 Education",
    "📣 Marketing / HR":              "📣 Marketing/HR",
    "🔬 Science / Research":          "🔬 Science",
    "🌐 General":                     "🌐 General",
}

def detect_domain(skills):
    domain_count = {}
    for skill in skills:
        d = ALL_SKILLS.get(skill.lower())
        if d and d != "🌐 General":
            domain_count[d] = domain_count.get(d, 0) + 1
    if not domain_count:
        return "🌐 General"
    best_short = max(domain_count, key=domain_count.get)
    for full, short in DOMAIN_MAP.items():
        if short == best_short:
            return full
    return best_short

# ============================================================
# 📄 EXTRACT TEXT
# ============================================================
def extract_text(filepath):
    text = ""
    try:
        if filepath.endswith(".pdf"):
            with open(filepath, "rb") as f:
                reader = PyPDF2.PdfReader(f)
                for page in reader.pages:
                    content = page.extract_text()
                    if content:
                        text += content + "\n"
        elif filepath.endswith(".docx"):
            d = docx.Document(filepath)
            for para in d.paragraphs:
                text += para.text + "\n"
    except Exception as e:
        print("Extract Error:", e)
    return text

# ============================================================
# 👤 EXTRACT BASIC INFO
# ============================================================
LOCATION_WORDS = {
    "sirsa","haryana","delhi","mumbai","bangalore","bengaluru","chennai",
    "kolkata","hyderabad","pune","jaipur","lucknow","chandigarh","amritsar",
    "ludhiana","gurugram","noida","agra","varanasi","patna","bhopal","indore",
    "nagpur","surat","ahmedabad","rajasthan","punjab","gujarat","maharashtra",
    "kerala","karnataka","uttar","pradesh","madhya","andhra","telangana",
    "bihar","jharkhand","odisha","assam","west","bengal","himachal","gurgaon",
    "faridabad","meerut","nashik","aurangabad","solapur","ranchi","raipur",
    "coimbatore","visakhapatnam","madurai","vijayawada","mysore","jodhpur",
    "india","new","north","south","east","west","central",
    "page","pages","resume","cv","curriculum","vitae","document","report",
    "candidate","applicant","profile","scanned","scanner","oken",
    "to","the","a","an","and","or","of","in","at","by","for","with",
    "is","are","was","were","be","been","has","have","had","do","does",
    "this","that","these","those","my","your","our","their","its",
    "mr","mrs","ms","dr","prof","sir","dear","from","subject","re",
    "am","i","we","they","he","she","it","who","which","what",
    "career","objective","skills","education","experience","projects",
    "languages","summary","contact","email","phone","address","mobile",
    "references","internship","bachelor","master","computer","applications",
    "ongoing","expected","completion","developer","frontend","backend","intern",
    "about","me","declaration","hobbies","interests","achievements",
    "certifications","workshops","seminars","activities","personal","details",
    "date","birth","gender","nationality","marital","status","male","female",
    "native","elementary","intermediate","fluent","proficient","english",
    "hindi","punjabi","training","activity","volunteer","nss","university",
    "student","disciplined","dedicated","skilled","motivated","passionate",
    "information","personal","curriculum","vitae2",
}

def extract_basic_info(text):
    clean = text.replace("\r", " ")
    email_match = re.search(
        r"[a-zA-Z0-9][a-zA-Z0-9+_.%-]*@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}", clean)
    email = email_match.group(0) if email_match else "Not Found"
    phone_match = re.search(
        r"(\+91[\s\-]?\d{5}[\s\-]?\d{5}|\+91[\s\-]?\d{10}|\b\d{5}[\s\-]\d{5}\b|\b[6-9]\d{9}\b)",
        clean)
    phone = phone_match.group(0).strip() if phone_match else "Not Found"
    name  = _detect_name(text)
    return name, phone, email

def _clean_line(line):
    line = re.sub(r'[|,/\\@#*\xb7:;_=+<>{}()\[\]"\'\`]', ' ', line)
    line = re.sub(r'\s+', ' ', line).strip()
    return line

def _is_valid_name_word(w):
    return w.isalpha() and len(w) >= 2 and w.lower() not in LOCATION_WORDS

def _extract_name_from_line(line, max_words=4):
    line_c = _clean_line(line)
    if not line_c: return None
    if re.search(r'\d', line_c): return None
    if len(line_c) > 50: return None
    words = line_c.split()
    valid = [w for w in words if _is_valid_name_word(w)]
    if not (1 <= len(valid) <= max_words): return None
    candidate = []
    for w in valid:
        if w[0].isupper() or w.isupper():
            candidate.append(w.title())
        else:
            break
    if not (1 <= len(candidate) <= max_words): return None
    if ' '.join(candidate).lower() in LOCATION_WORDS: return None
    return ' '.join(candidate)

def _detect_name(text):
    lines = [l.strip() for l in text.split('\n') if l.strip()]
    merged_lines = []
    i = 0
    while i < len(lines):
        line  = lines[i]
        lc    = _clean_line(line)
        words = lc.split()
        if (len(words) == 1 and words[0].isupper() and words[0].isalpha()
                and len(words[0]) >= 2 and words[0].lower() not in LOCATION_WORDS):
            if i + 1 < len(lines):
                next_lc    = _clean_line(lines[i+1])
                next_words = next_lc.split()
                if (len(next_words) == 1 and next_words[0].isupper()
                        and next_words[0].isalpha() and len(next_words[0]) >= 2
                        and next_words[0].lower() not in LOCATION_WORDS):
                    merged_lines.append(words[0] + ' ' + next_words[0])
                    i += 2
                    continue
        merged_lines.append(line)
        i += 1
    all_lines = merged_lines + lines
    for line in all_lines[:12]:
        if re.match(r'^page\s*\d*$', line.strip().lower()): continue
        name = _extract_name_from_line(line, max_words=4)
        if name: return name
    for line in lines[:15]:
        if re.match(r'^page\s*\d*$', line.strip().lower()): continue
        lc = _clean_line(line)
        if not lc or re.search(r'\d', lc) or len(lc) > 45: continue
        words = lc.split()
        if (len(words) >= 1 and all(w.isupper() and w.isalpha() and len(w) >= 2 for w in words)):
            valid = [w.title() for w in words if w.lower() not in LOCATION_WORDS]
            if 1 <= len(valid) <= 4: return ' '.join(valid)
    label_pat = re.compile(
        r'(?:name|full\s*name|candidate\s*name)\s*[:\-]\s*([A-Za-z][A-Za-z\s]{1,40})',
        re.IGNORECASE)
    lm = label_pat.search(text)
    if lm:
        raw   = lm.group(1).strip()
        words = raw.split()
        valid = [w.title() for w in words
                 if w.isalpha() and len(w) >= 2
                 and w.lower() not in LOCATION_WORDS
                 and w.lower() not in ('name','full','candidate')]
        if valid: return ' '.join(valid[:3])
    return 'Not Found'

# ============================================================
# ⚡ SKILL DETECTION
# ============================================================
STRICT_BOUNDARY = {
    "sql","git","aws","seo","bim","plc","icu","ecg","bds","llb",
    "cfa","cpa","sap","css","php","vue","law","ca",
}

def detect_skills_ai(text):
    text_lower = text.lower()
    found = []
    for skill in ALL_SKILLS.keys():
        if skill in STRICT_BOUNDARY or len(skill) <= 3:
            pattern = r'\b' + re.escape(skill) + r'\b'
            if re.search(pattern, text_lower):
                found.append(skill)
        else:
            if skill in text_lower:
                found.append(skill)
    seen, unique = set(), []
    for s in found:
        if s not in seen:
            seen.add(s)
            unique.append(s)
    return unique[:25]

# ============================================================
# ⭐ SCORE & GRADE
# ============================================================
def calculate_score(skills):
    c = len(skills)
    if c >= 15:   return 10
    elif c >= 12: return 9
    elif c >= 10: return 8
    elif c >= 8:  return 7
    elif c >= 6:  return 6
    elif c >= 4:  return 5
    elif c >= 2:  return 4
    else:         return 3

def calculate_grade(score):
    if score >= 9:   return "A+"
    elif score >= 8: return "A"
    elif score >= 7: return "B+"
    elif score >= 6: return "B"
    elif score >= 5: return "C"
    else:            return "D"

# ============================================================
# 💼 JOB MATCHING
# ============================================================
DOMAIN_JOB_FILTER = {
    "💻 IT":           ["Frontend Developer","Backend Developer","Full Stack Developer",
                        "Java Developer","DevOps Engineer","Cloud Engineer",
                        "Android Developer","Software Tester / QA","UI/UX Designer"],
    "📊 Data Science": ["Data Analyst","Data Scientist","ML Engineer","AI Engineer",
                        "Business Intelligence","Data Engineer","NLP Engineer",
                        "Computer Vision Engineer","Statistician"],
    "🏥 Medical":      ["Medical Doctor (MBBS)","Nurse","Dentist (BDS)","Pharmacist",
                        "Medical Coder","Physiotherapist","Clinical Research Assoc.",
                        "Bioinformatics Analyst"],
    "⚖️ Law":          ["Corporate Lawyer","Criminal Lawyer","Law Intern","Legal Advisor",
                        "Tax Consultant","Paralegal","IP Lawyer","Cyber Law Specialist",
                        "Family Lawyer"],
    "📈 Finance":      ["Chartered Accountant","Financial Analyst","Accountant",
                        "Tax Analyst","Audit Associate","Stock Market Analyst"],
    "🎨 Design":       ["UI/UX Designer","Graphic Designer","Video Editor",
                        "Brand Designer","Motion Designer","Product Designer"],
    "🏗️ Engineering": ["Civil Engineer","Mechanical Engineer","Structural Engineer",
                        "VLSI Engineer"],
    "📚 Education":    ["School Teacher","Corporate Trainer","Online Tutor",
                        "Curriculum Designer"],
    "📣 Marketing/HR": ["Digital Marketing Manager","HR Manager","Content Writer",
                        "Social Media Manager","Recruitment Specialist","HR Business Partner"],
    "🔬 Science":      ["Research Scientist","Biotechnologist","Microbiologist",
                        "Lab Technician","Clinical Research Assoc.","Bioinformatics Analyst"],
    "🌐 General":      [],
}

def match_jobs(skills, domain="🌐 General"):
    skills_lower = [s.lower() for s in skills]
    domain_lower = domain.lower()
    allowed_jobs = None
    if "law" in domain_lower:            allowed_jobs = DOMAIN_JOB_FILTER["⚖️ Law"]
    elif "finance" in domain_lower:      allowed_jobs = DOMAIN_JOB_FILTER["📈 Finance"]
    elif "medical" in domain_lower:      allowed_jobs = DOMAIN_JOB_FILTER["🏥 Medical"]
    elif "data science" in domain_lower: allowed_jobs = DOMAIN_JOB_FILTER["📊 Data Science"]
    elif "design" in domain_lower:       allowed_jobs = DOMAIN_JOB_FILTER["🎨 Design"]
    elif "engineering" in domain_lower:  allowed_jobs = DOMAIN_JOB_FILTER["🏗️ Engineering"]
    elif "education" in domain_lower:    allowed_jobs = DOMAIN_JOB_FILTER["📚 Education"]
    elif "marketing" in domain_lower:    allowed_jobs = DOMAIN_JOB_FILTER["📣 Marketing/HR"]
    elif "science" in domain_lower:      allowed_jobs = DOMAIN_JOB_FILTER["🔬 Science"]
    elif "it" in domain_lower or "web" in domain_lower: allowed_jobs = DOMAIN_JOB_FILTER["💻 IT"]
    else:                                allowed_jobs = []
    results = []
    for job in JOB_DATABASE:
        if allowed_jobs is not None and len(allowed_jobs) > 0:
            if job["job"] not in allowed_jobs: continue
        required = job["required"]
        matched  = [r for r in required if any(r in s for s in skills_lower)]
        missing  = [r for r in required if r not in matched]
        pct      = int((len(matched) / len(required)) * 100) if required else 0
        results.append({"job": job["job"], "match": pct, "required": required,
                         "missing": missing, "is_full": (pct == 100)})
    results.sort(key=lambda x: x["match"], reverse=True)
    return results

# ============================================================
# 💡 SUGGESTIONS
# ============================================================
DOMAIN_TIPS = {
    "💻 IT / Web Development": {
        "platform": "freeCodeCamp / Udemy / LeetCode",
        "tips": ["Build 2-3 portfolio projects and put them on GitHub",
                 "Practice DSA on LeetCode or HackerRank daily",
                 "Get AWS Cloud Practitioner certification",
                 "Contribute to open-source projects",
                 "Add your GitHub profile link on resume"],
    },
    "📊 Data Science / ML / AI": {
        "platform": "Kaggle / Coursera / fast.ai",
        "tips": ["Participate in a Kaggle competition",
                 "Build end-to-end ML projects on GitHub",
                 "Get Google Data Analytics or IBM Data Science certificate",
                 "Learn SQL deeply — it is essential for all data roles",
                 "Create a project using real datasets (UCI, Kaggle)"],
    },
    "🏥 Medical / Healthcare": {
        "platform": "NMC Guidelines / PubMed / USMLE",
        "tips": ["Complete clinical internship hours",
                 "Learn medical coding (ICD-10, CPT codes)",
                 "Stay updated with latest clinical research",
                 "Add any specialization or certification to resume",
                 "Mention any research papers or case studies"],
    },
    "⚖️ Law / Legal": {
        "platform": "SCC Online / Manupatra / LexisNexis",
        "tips": ["Participate in moot court competitions",
                 "Do internship at a reputed law firm",
                 "Learn legal research tools (Manupatra, SCC Online)",
                 "Specialize in a niche area of law (criminal, corporate, cyber)",
                 "Build strong legal writing and contract drafting skills"],
    },
    "📈 Finance / Accounting": {
        "platform": "ICAI / NSE Academy / Coursera",
        "tips": ["Pursue CA / CFA / CPA certification",
                 "Master advanced Excel and financial modeling",
                 "Stay updated with latest GST and income tax laws",
                 "Practice with Tally ERP and SAP",
                 "Add any internship or articleship experience"],
    },
    "🎨 Design / Creative": {
        "platform": "Behance / Dribbble / Adobe Learn",
        "tips": ["Build a strong Behance or Dribbble portfolio",
                 "Learn Figma for modern UI/UX design",
                 "Study color theory and typography fundamentals",
                 "Do freelance projects to build real experience",
                 "Add your portfolio link in your resume"],
    },
    "🏗️ Civil / Mech. Engineering": {
        "platform": "NPTEL / Coursera / AutoDesk Learn",
        "tips": ["Get certified in AutoCAD, Revit, or STAAD Pro",
                 "Do internship with construction or manufacturing firm",
                 "Learn BIM software for civil engineering roles",
                 "Prepare for GATE if targeting PSU jobs"],
    },
    "📚 Education / Teaching": {
        "platform": "DIKSHA / Coursera / NCERT Portal",
        "tips": ["Get B.Ed or CTET certification",
                 "Learn e-learning platforms (Moodle, Google Classroom)",
                 "Build engaging and structured lesson plans",
                 "Take public speaking or communication courses"],
    },
    "📣 Marketing / HR": {
        "platform": "Google Digital Garage / HubSpot / LinkedIn Learning",
        "tips": ["Get Google Analytics or HubSpot certification (free)",
                 "Build your personal social media brand",
                 "Learn basic SEO and content strategy",
                 "Mention campaign results with numbers (e.g. 30% growth)"],
    },
    "🔬 Science / Research": {
        "platform": "PubMed / ResearchGate / Coursera",
        "tips": ["Publish research in indexed journals",
                 "Learn data analysis tools for research (SPSS, R, Python)",
                 "Apply for funded research or fellowship programs",
                 "Mention lab instruments and techniques you know"],
    },
}

def generate_suggestions(skills, missing_skills, domain):
    lines = [f"## Domain Detected: {domain}\n"]
    if missing_skills:
        lines.append("## Skills to Learn:")
        for s in list(missing_skills)[:7]:
            lines.append(f"- {s}")
    matched_tips = None
    for key in DOMAIN_TIPS:
        if key in domain:
            matched_tips = DOMAIN_TIPS[key]
            break
    if matched_tips:
        lines.append(f"\n## Recommended Platforms:")
        lines.append(f"- {matched_tips['platform']}")
        lines.append("\n## Career Tips:")
        for tip in matched_tips["tips"]:
            lines.append(f"- {tip}")
    lines += ["\n## General Resume Tips:",
              "- Add a strong professional summary at the top",
              "- Quantify your achievements (e.g. improved performance by 30%)",
              "- Keep resume clean and to 1-2 pages only",
              "- Add LinkedIn profile and GitHub or portfolio link",
              "- Use action verbs: Built, Designed, Managed, Led, Developed",
              "- Proofread for spelling and grammar mistakes"]
    return "\n".join(lines)

# ============================================================
# 📄 ANALYZER PDF REPORT  (unchanged)
# ============================================================
def create_pdf(score, grade, skills, suggestion,
               name="N/A", phone="N/A", email="N/A", domain="General"):
    doc = SimpleDocTemplate("report.pdf",
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=12*mm, bottomMargin=12*mm)
    DARK   = colors.HexColor("#0f172a")
    CYAN   = colors.HexColor("#06b6d4")
    BLUE   = colors.HexColor("#3b82f6")
    GREEN  = colors.HexColor("#22c55e")
    LIGHT  = colors.HexColor("#f0f9ff")
    WHITE  = colors.white
    GREY   = colors.HexColor("#64748b")
    SILVER = colors.HexColor("#e2e8f0")
    title_style = ParagraphStyle("ts", fontSize=22, fontName="Helvetica-Bold",
        textColor=WHITE, alignment=TA_CENTER, spaceAfter=4)
    sub_style   = ParagraphStyle("ss", fontSize=10, fontName="Helvetica",
        textColor=colors.HexColor("#94a3b8"), alignment=TA_CENTER)
    sec_style   = ParagraphStyle("secs", fontSize=12, fontName="Helvetica-Bold",
        textColor=CYAN, spaceBefore=10, spaceAfter=4)
    lbl_style   = ParagraphStyle("lbl", fontSize=9, fontName="Helvetica-Bold", textColor=GREY)
    val_style   = ParagraphStyle("val", fontSize=10, fontName="Helvetica", textColor=DARK)
    body_style  = ParagraphStyle("body", fontSize=9, fontName="Helvetica", textColor=DARK, leading=14)
    head_style  = ParagraphStyle("head", fontSize=10, fontName="Helvetica-Bold",
        textColor=BLUE, spaceBefore=8, spaceAfter=3)
    blt_style   = ParagraphStyle("blt", fontSize=9, fontName="Helvetica",
        textColor=colors.HexColor("#1e293b"), leading=14, leftIndent=10)
    story = []
    story.append(Table([[Paragraph("AI Resume Analysis Report", title_style)]], colWidths=[180*mm]))
    story[-1].setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),DARK),
        ("TOPPADDING",(0,0),(-1,-1),14),("BOTTOMPADDING",(0,0),(-1,-1),8),("LEFTPADDING",(0,0),(-1,-1),10)]))
    date_str = datetime.now().strftime("%d %B %Y  |  %I:%M %p")
    story.append(Table([[Paragraph(f"Generated on: {date_str}", sub_style)]], colWidths=[180*mm]))
    story[-1].setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#1e293b")),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),8)]))
    story.append(Spacer(1,8))
    story.append(Paragraph("Candidate Information", sec_style))
    story.append(HRFlowable(width="100%", thickness=1, color=CYAN, spaceAfter=6))
    info_rows = [
        [Paragraph("Name",lbl_style),  Paragraph(str(name),val_style),
         Paragraph("Phone",lbl_style), Paragraph(str(phone),val_style)],
        [Paragraph("Email",lbl_style), Paragraph(str(email),val_style),
         Paragraph("Domain",lbl_style),Paragraph(str(domain),val_style)],
    ]
    t = Table(info_rows, colWidths=[22*mm,68*mm,22*mm,68*mm])
    t.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),LIGHT),("GRID",(0,0),(-1,-1),0.4,SILVER),
        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),("LEFTPADDING",(0,0),(-1,-1),8)]))
    story.append(t); story.append(Spacer(1,10))
    story.append(Paragraph("Score & Grade", sec_style))
    story.append(HRFlowable(width="100%", thickness=1, color=CYAN, spaceAfter=6))
    grade_color = (GREEN if score>=8 else BLUE if score>=6
                   else colors.HexColor("#f59e0b") if score>=4 else colors.HexColor("#ef4444"))
    sc_lbl = ParagraphStyle("sc",fontSize=20,fontName="Helvetica-Bold",textColor=WHITE,alignment=TA_CENTER)
    gr_lbl = ParagraphStyle("gr",fontSize=20,fontName="Helvetica-Bold",textColor=WHITE,alignment=TA_CENTER)
    sc_sub = ParagraphStyle("scsub",fontSize=9,fontName="Helvetica",
                             textColor=colors.HexColor("#94a3b8"),alignment=TA_CENTER)
    sd = [[Table([[Paragraph(f"{score} / 10",sc_lbl)],[Paragraph("Resume Score",sc_sub)]],colWidths=[85*mm]),
           Table([[Paragraph(f"  {grade}  ",gr_lbl)],[Paragraph("Grade",sc_sub)]],colWidths=[85*mm])]]
    st = Table(sd, colWidths=[90*mm,90*mm])
    st.setStyle(TableStyle([("BACKGROUND",(0,0),(0,0),BLUE),("BACKGROUND",(1,0),(1,0),grade_color),
        ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(-1,-1),6),("INNERGRID",(0,0),(-1,-1),0,WHITE)]))
    story.append(st); story.append(Spacer(1,10))
    story.append(Paragraph("Detected Skills", sec_style))
    story.append(HRFlowable(width="100%", thickness=1, color=CYAN, spaceAfter=6))
    if skills:
        sk_style = ParagraphStyle("sk",fontSize=9,fontName="Helvetica",textColor=colors.HexColor("#0369a1"))
        cols, rows = 3, []
        for i in range(0, len(skills), cols):
            row = []
            for j in range(cols):
                idx = i+j
                row.append(Paragraph(f"✦  {skills[idx].title()}",sk_style)
                           if idx<len(skills) else Paragraph("",sk_style))
            rows.append(row)
        col_w = 180*mm/cols
        kt = Table(rows, colWidths=[col_w]*cols)
        kt.setStyle(TableStyle([("BACKGROUND",(0,0),(-1,-1),LIGHT),("GRID",(0,0),(-1,-1),0.3,SILVER),
            ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),("LEFTPADDING",(0,0),(-1,-1),8)]))
        story.append(kt)
    else:
        story.append(Paragraph("No skills detected.", body_style))
    story.append(Spacer(1,10))
    story.append(Paragraph("Suggestions & Recommendations", sec_style))
    story.append(HRFlowable(width="100%", thickness=1, color=CYAN, spaceAfter=6))
    for line in suggestion.split("\n"):
        line = line.strip()
        if not line:               story.append(Spacer(1,3))
        elif line.startswith("##"):story.append(Paragraph(f"▸  {line.replace('##','').strip()}", head_style))
        elif line.startswith("-"): story.append(Paragraph(f"  •  {line[1:].strip()}", blt_style))
        else:                      story.append(Paragraph(line, body_style))
    story.append(Spacer(1,14))
    story.append(HRFlowable(width="100%", thickness=1, color=SILVER))
    ft_style = ParagraphStyle("ft",fontSize=8,fontName="Helvetica",textColor=GREY,alignment=TA_CENTER,spaceBefore=6)
    story.append(Paragraph("Generated by AI Resume Analyzer  •  Developed by Kirandeep Kaur  •  Confidential", ft_style))
    doc.build(story)
    print("Analyzer PDF created ✅")

# ============================================================
# 📄 BUILDER PDF — Beautiful Professional Resume
# ============================================================
def create_resume_pdf(data, filepath="built_resume.pdf"):
    doc = SimpleDocTemplate(filepath,
        leftMargin=14*mm, rightMargin=14*mm,
        topMargin=0*mm,   bottomMargin=10*mm)

    # ── Color Palette ─────────────────────────────────────────
    NAVY      = colors.HexColor("#1a237e")   # dark header
    INDIGO    = colors.HexColor("#283593")   # sidebar / section bar
    STEEL     = colors.HexColor("#3949ab")   # accent lines
    LIGHT_BG  = colors.HexColor("#f0f4ff")   # table row background
    PALE      = colors.HexColor("#e8eaf6")   # table header row
    DARK_TEXT = colors.HexColor("#1a1a2e")
    GREY_TEXT = colors.HexColor("#555577")
    WHITE     = colors.white
    SILVER    = colors.HexColor("#c5cae9")
    GREEN_ACC = colors.HexColor("#00897b")

    # ── Styles ────────────────────────────────────────────────
    name_s = ParagraphStyle("name_s", fontSize=28, fontName="Helvetica-Bold",
        textColor=WHITE, alignment=TA_LEFT, leading=32)
    role_s = ParagraphStyle("role_s", fontSize=12, fontName="Helvetica",
        textColor=colors.HexColor("#c5cae9"), alignment=TA_LEFT, leading=16)
    contact_s = ParagraphStyle("contact_s", fontSize=8.5, fontName="Helvetica",
        textColor=colors.HexColor("#e8eaf6"), alignment=TA_LEFT, leading=13)

    sec_s = ParagraphStyle("sec_s", fontSize=10, fontName="Helvetica-Bold",
        textColor=NAVY, spaceBefore=8, spaceAfter=2, leading=14,
        borderPad=2)
    body_s = ParagraphStyle("body_s", fontSize=9, fontName="Helvetica",
        textColor=DARK_TEXT, leading=13)
    bold_s = ParagraphStyle("bold_s", fontSize=9, fontName="Helvetica-Bold",
        textColor=DARK_TEXT, leading=13)
    small_s = ParagraphStyle("small_s", fontSize=8, fontName="Helvetica",
        textColor=GREY_TEXT, leading=12)
    bullet_s = ParagraphStyle("blt_s", fontSize=9, fontName="Helvetica",
        textColor=DARK_TEXT, leading=13, leftIndent=12)
    skill_s = ParagraphStyle("skill_s", fontSize=8.5, fontName="Helvetica",
        textColor=NAVY, alignment=TA_CENTER)
    tbl_hdr_s = ParagraphStyle("tbl_hdr", fontSize=8.5, fontName="Helvetica-Bold",
        textColor=WHITE, alignment=TA_CENTER)
    tbl_cell_s = ParagraphStyle("tbl_cell", fontSize=8.5, fontName="Helvetica",
        textColor=DARK_TEXT, alignment=TA_CENTER, leading=12)
    tbl_left_s = ParagraphStyle("tbl_left", fontSize=8.5, fontName="Helvetica",
        textColor=DARK_TEXT, alignment=TA_LEFT, leading=12)

    story = []

    # ══════════════════════════════════════════════════════════
    # HEADER BANNER  (full-width dark navy)
    # ══════════════════════════════════════════════════════════
    name      = data.get("name","") or "Your Name"
    job_title = data.get("job_title","") or ""
    email     = data.get("email","") or ""
    phone     = data.get("phone","") or ""
    location  = data.get("location","") or ""
    linkedin  = data.get("linkedin","") or ""

    contact_parts = [x for x in [email, phone, location, linkedin] if x]
    contact_line  = "   |   ".join(contact_parts)

    # Left side: name + role
    left_cell = [Paragraph(name, name_s)]
    if job_title:
        left_cell.append(Paragraph(job_title, role_s))
    if contact_line:
        left_cell.append(Spacer(1, 4))
        left_cell.append(Paragraph(contact_line, contact_s))

    # Right side: date generated
    date_s = ParagraphStyle("date_s", fontSize=8, fontName="Helvetica",
        textColor=colors.HexColor("#9fa8da"), alignment=TA_RIGHT)
    right_cell = [Paragraph(datetime.now().strftime("%B %Y"), date_s)]

    header_tbl = Table([[left_cell, right_cell]],
                       colWidths=[145*mm, 37*mm])
    header_tbl.setStyle(TableStyle([
        ("BACKGROUND",    (0,0), (-1,-1), NAVY),
        ("TOPPADDING",    (0,0), (-1,-1), 16),
        ("BOTTOMPADDING", (0,0), (-1,-1), 16),
        ("LEFTPADDING",   (0,0), (0,0),  14),
        ("RIGHTPADDING",  (1,0), (1,0),  14),
        ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
    ]))
    story.append(header_tbl)
    story.append(Spacer(1, 10))

    # ── Helper: section heading with colored left bar ─────────
    def sec_header(title):
        bar_tbl = Table([[Paragraph(f"  {title}", sec_s)]], colWidths=[182*mm])
        bar_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), PALE),
            ("LEFTPADDING",   (0,0), (-1,-1), 0),
            ("RIGHTPADDING",  (0,0), (-1,-1), 6),
            ("TOPPADDING",    (0,0), (-1,-1), 4),
            ("BOTTOMPADDING", (0,0), (-1,-1), 4),
            ("LINEBELOW",     (0,0), (-1,-1), 1.5, STEEL),
            ("LINEBEFORE",    (0,0), (0,-1),  4,   STEEL),
        ]))
        story.append(bar_tbl)
        story.append(Spacer(1, 4))

    # ══════════════════════════════════════════════════════════
    # PROFESSIONAL SUMMARY
    # ══════════════════════════════════════════════════════════
    summary = data.get("summary","").strip()
    if summary:
        sec_header("PROFESSIONAL SUMMARY")
        summary_tbl = Table([[Paragraph(summary, body_s)]], colWidths=[182*mm])
        summary_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), LIGHT_BG),
            ("LEFTPADDING",   (0,0), (-1,-1), 10),
            ("RIGHTPADDING",  (0,0), (-1,-1), 10),
            ("TOPPADDING",    (0,0), (-1,-1), 8),
            ("BOTTOMPADDING", (0,0), (-1,-1), 8),
            ("LINEAFTER",     (0,0), (0,-1),  0.5, SILVER),
        ]))
        story.append(summary_tbl)
        story.append(Spacer(1, 8))

    # ══════════════════════════════════════════════════════════
    # SKILLS — pill-style badge table
    # ══════════════════════════════════════════════════════════
    skills = data.get("skills", [])
    if skills:
        sec_header("SKILLS")
        COLS = 5
        rows = []
        for i in range(0, len(skills), COLS):
            row = []
            for j in range(COLS):
                idx = i + j
                if idx < len(skills):
                    row.append(Paragraph(skills[idx].title(), skill_s))
                else:
                    row.append(Paragraph("", skill_s))
            rows.append(row)
        col_w = 182*mm / COLS
        sk_tbl = Table(rows, colWidths=[col_w]*COLS)
        sk_tbl.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), LIGHT_BG),
            ("GRID",          (0,0), (-1,-1), 0.5, SILVER),
            ("TOPPADDING",    (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("ALIGN",         (0,0), (-1,-1), "CENTER"),
            ("ROWBACKGROUNDS",(0,0), (-1,-1), [LIGHT_BG, WHITE]),
        ]))
        story.append(sk_tbl)
        story.append(Spacer(1, 8))

    # ══════════════════════════════════════════════════════════
    # WORK EXPERIENCE
    # ══════════════════════════════════════════════════════════
    experience = data.get("experience", [])
    exp_valid  = [e for e in experience if e.get("title") or e.get("company")]
    if exp_valid:
        sec_header("WORK EXPERIENCE")
        for idx, exp in enumerate(exp_valid):
            title    = exp.get("title","")
            company  = exp.get("company","")
            duration = exp.get("duration","")
            loc      = exp.get("location","")
            desc     = exp.get("desc","").strip()
            right_t  = "  |  ".join([x for x in [duration, loc] if x])

            # Title row
            exp_hdr = Table([[
                Paragraph(f"<b>{title}</b>", bold_s),
                Paragraph(right_t, small_s)
            ]], colWidths=[130*mm, 52*mm])
            exp_hdr.setStyle(TableStyle([
                ("BACKGROUND",    (0,0), (-1,-1), PALE),
                ("TOPPADDING",    (0,0), (-1,-1), 5),
                ("BOTTOMPADDING", (0,0), (-1,-1), 5),
                ("LEFTPADDING",   (0,0), (0,0),  10),
                ("RIGHTPADDING",  (1,0), (1,0),  8),
                ("ALIGN",         (1,0), (1,0),  "RIGHT"),
                ("VALIGN",        (0,0), (-1,-1), "MIDDLE"),
                ("LINEBEFORE",    (0,0), (0,-1),  3, STEEL),
            ]))
            story.append(exp_hdr)
            if company:
                story.append(Paragraph(f"  {company}", small_s))
            if desc:
                for line in desc.split("\n"):
                    line = line.strip()
                    if line:
                        bullet = line[1:].strip() if line.startswith("•") else line
                        story.append(Paragraph(f"  •  {bullet}", bullet_s))
            story.append(Spacer(1, 6))

    # ══════════════════════════════════════════════════════════
    # EDUCATION — Beautiful table with headers
    # ══════════════════════════════════════════════════════════
    education = data.get("education", [])
    edu_valid  = [e for e in education if e.get("degree") or e.get("institution")]
    if edu_valid:
        sec_header("EDUCATION")

        # Table headers
        edu_hdr_row = [
            Paragraph("Degree / Course", tbl_hdr_s),
            Paragraph("Institution",     tbl_hdr_s),
            Paragraph("Year",            tbl_hdr_s),
            Paragraph("Grade / CGPA",    tbl_hdr_s),
        ]
        edu_rows = [edu_hdr_row]
        for i, edu in enumerate(edu_valid):
            bg = LIGHT_BG if i % 2 == 0 else WHITE
            edu_rows.append([
                Paragraph(edu.get("degree",""),      tbl_left_s),
                Paragraph(edu.get("institution",""), tbl_left_s),
                Paragraph(edu.get("year",""),        tbl_cell_s),
                Paragraph(edu.get("grade",""),       tbl_cell_s),
            ])

        edu_tbl = Table(edu_rows, colWidths=[65*mm, 72*mm, 25*mm, 20*mm])
        # Build row backgrounds (alternating) after header
        row_bgs = []
        for i in range(1, len(edu_rows)):
            c = LIGHT_BG if i % 2 == 1 else WHITE
            row_bgs.append(("BACKGROUND", (0,i), (-1,i), c))

        edu_tbl.setStyle(TableStyle([
            # Header
            ("BACKGROUND",    (0,0), (-1,0), INDIGO),
            ("TEXTCOLOR",     (0,0), (-1,0), WHITE),
            ("FONTNAME",      (0,0), (-1,0), "Helvetica-Bold"),
            ("FONTSIZE",      (0,0), (-1,0), 8.5),
            ("ALIGN",         (0,0), (-1,0), "CENTER"),
            ("TOPPADDING",    (0,0), (-1,0), 7),
            ("BOTTOMPADDING", (0,0), (-1,0), 7),
            # Data rows
            ("FONTSIZE",      (0,1), (-1,-1), 8.5),
            ("TOPPADDING",    (0,1), (-1,-1), 6),
            ("BOTTOMPADDING", (0,1), (-1,-1), 6),
            ("LEFTPADDING",   (0,1), (-1,-1), 8),
            ("ALIGN",         (2,1), (-1,-1), "CENTER"),
            # Grid
            ("GRID",          (0,0), (-1,-1), 0.5, SILVER),
            ("LINEBELOW",     (0,0), (-1,0),  1.5, STEEL),
            ("LINEBEFORE",    (0,0), (0,-1),  3,   STEEL),
        ] + row_bgs))
        story.append(edu_tbl)
        story.append(Spacer(1, 8))

    # ══════════════════════════════════════════════════════════
    # PROJECTS
    # ══════════════════════════════════════════════════════════
    projects   = data.get("projects", [])
    proj_valid = [p for p in projects if p.get("title")]
    if proj_valid:
        sec_header("PROJECTS")
        for proj in proj_valid:
            title = proj.get("title","")
            tech  = proj.get("tech","")
            desc  = proj.get("desc","").strip()

            tech_badge = ""
            if tech:
                tech_badge = f'  <font color="#3949ab">[{tech}]</font>'

            proj_hdr = Table([[
                Paragraph(f"<b>{title}</b>{tech_badge}", bold_s)
            ]], colWidths=[182*mm])
            proj_hdr.setStyle(TableStyle([
                ("BACKGROUND",  (0,0), (-1,-1), PALE),
                ("LEFTPADDING", (0,0), (-1,-1), 10),
                ("TOPPADDING",  (0,0), (-1,-1), 5),
                ("BOTTOMPADDING",(0,0),(-1,-1), 5),
                ("LINEBEFORE",  (0,0), (0,-1),  3, GREEN_ACC),
            ]))
            story.append(proj_hdr)
            if desc:
                story.append(Paragraph(f"  {desc}", body_s))
            story.append(Spacer(1, 6))

    # ══════════════════════════════════════════════════════════
    # CERTIFICATIONS
    # ══════════════════════════════════════════════════════════
    certs      = data.get("certifications", [])
    cert_valid = [c for c in certs if c.get("title")]
    if cert_valid:
        sec_header("CERTIFICATIONS & ACHIEVEMENTS")
        cert_rows = []
        for i, cert in enumerate(cert_valid):
            bg = LIGHT_BG if i % 2 == 0 else WHITE
            cert_rows.append([
                Paragraph(f"  ★  {cert.get('title','')}", body_s),
                Paragraph(cert.get("issuer",""), small_s),
            ])
        ct = Table(cert_rows, colWidths=[130*mm, 52*mm])
        bg_styles = [("BACKGROUND",(0,i),(-1,i), LIGHT_BG if i%2==0 else WHITE)
                     for i in range(len(cert_rows))]
        ct.setStyle(TableStyle([
            ("GRID",          (0,0), (-1,-1), 0.4, SILVER),
            ("TOPPADDING",    (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("LEFTPADDING",   (0,0), (0,-1),  8),
            ("ALIGN",         (1,0), (1,-1),  "RIGHT"),
            ("RIGHTPADDING",  (1,0), (1,-1),  8),
            ("LINEBEFORE",    (0,0), (0,-1),  3, GREEN_ACC),
        ] + bg_styles))
        story.append(ct)
        story.append(Spacer(1, 8))

    # ══════════════════════════════════════════════════════════
    # LANGUAGES & INTERESTS
    # ══════════════════════════════════════════════════════════
    languages = data.get("languages","").strip()
    hobbies   = data.get("hobbies","").strip()
    if languages or hobbies:
        sec_header("ADDITIONAL INFORMATION")
        add_rows = []
        if languages:
            add_rows.append([Paragraph("<b>Languages</b>", bold_s),
                             Paragraph(languages, body_s)])
        if hobbies:
            add_rows.append([Paragraph("<b>Interests</b>", bold_s),
                             Paragraph(hobbies, body_s)])
        at = Table(add_rows, colWidths=[35*mm, 147*mm])
        at.setStyle(TableStyle([
            ("BACKGROUND",    (0,0), (-1,-1), LIGHT_BG),
            ("TOPPADDING",    (0,0), (-1,-1), 6),
            ("BOTTOMPADDING", (0,0), (-1,-1), 6),
            ("LEFTPADDING",   (0,0), (0,-1),  10),
            ("GRID",          (0,0), (-1,-1), 0.4, SILVER),
            ("LINEBEFORE",    (0,0), (0,-1),  3, STEEL),
        ]))
        story.append(at)

    # ══════════════════════════════════════════════════════════
    # FOOTER
    # ══════════════════════════════════════════════════════════
    story.append(Spacer(1, 12))
    story.append(HRFlowable(width="100%", thickness=1, color=STEEL))
    ft = ParagraphStyle("ft", fontSize=7, fontName="Helvetica",
                        textColor=GREY_TEXT, alignment=TA_CENTER, spaceBefore=4)
    story.append(Paragraph(
        f"Resume of {name}  •  Built with Resume Builder  •  {datetime.now().strftime('%d %B %Y')}",
        ft))

    doc.build(story)
    print("Beautiful Resume PDF created ✅")

# ============================================================
# 🌐 FLASK ROUTES — ANALYZER  (completely unchanged)
# ============================================================
@app.route("/")
def home():
    return render_template("index.html")

@app.route("/upload", methods=["POST"])
def upload():
    print("UPLOAD HIT 🚀")
    file = request.files["resume"]
    path = os.path.join(app.config["UPLOAD_FOLDER"], file.filename)
    file.save(path)

    text               = extract_text(path)
    name, phone, email = extract_basic_info(text)
    skills             = detect_skills_ai(text)
    score              = calculate_score(skills)
    grade              = calculate_grade(score)
    domain             = detect_domain(skills)

    job_results = match_jobs(skills, domain)
    top_jobs    = job_results[:3]
    other_jobs  = job_results[3:9]

    all_missing = set()
    for j in job_results[:5]:
        for s in j["missing"]:
            all_missing.add(s)

    suggestion = generate_suggestions(skills, list(all_missing), domain)
    create_pdf(score, grade, skills, suggestion,
               name=name, phone=phone, email=email, domain=domain)

    return render_template("result.html",
        top_jobs=top_jobs, other_jobs=other_jobs,
        user_skills=skills, mapped_skills={},
        score=score, grade=grade, suggestion=suggestion,
        domain=domain, name=name, phone=phone, email=email)

@app.route("/download")
def download():
    return send_file("report.pdf", as_attachment=True)

# ============================================================
# 🌐 FLASK ROUTES — BUILDER
# ============================================================
@app.route("/builder")
def builder():
    return render_template("builder.html", error=None)


@app.route("/build-resume", methods=["POST"])
def build_resume():
    f = request.form

    skills_raw = f.get("skills", "")
    skills     = [s.strip() for s in skills_raw.split(",") if s.strip()]

    edu_degrees      = f.getlist("edu_degree[]")
    edu_institutions = f.getlist("edu_institution[]")
    edu_years        = f.getlist("edu_year[]")
    edu_grades       = f.getlist("edu_grade[]")
    education = []
    for i in range(len(edu_degrees)):
        education.append({
            "degree":      edu_degrees[i]      if i < len(edu_degrees)      else "",
            "institution": edu_institutions[i] if i < len(edu_institutions) else "",
            "year":        edu_years[i]        if i < len(edu_years)        else "",
            "grade":       edu_grades[i]       if i < len(edu_grades)       else "",
        })

    exp_titles    = f.getlist("exp_title[]")
    exp_companies = f.getlist("exp_company[]")
    exp_durations = f.getlist("exp_duration[]")
    exp_locations = f.getlist("exp_location[]")
    exp_descs     = f.getlist("exp_desc[]")
    experience = []
    for i in range(len(exp_titles)):
        experience.append({
            "title":    exp_titles[i]    if i < len(exp_titles)    else "",
            "company":  exp_companies[i] if i < len(exp_companies) else "",
            "duration": exp_durations[i] if i < len(exp_durations) else "",
            "location": exp_locations[i] if i < len(exp_locations) else "",
            "desc":     exp_descs[i]     if i < len(exp_descs)     else "",
        })

    proj_titles = f.getlist("proj_title[]")
    proj_techs  = f.getlist("proj_tech[]")
    proj_descs  = f.getlist("proj_desc[]")
    projects = []
    for i in range(len(proj_titles)):
        projects.append({
            "title": proj_titles[i] if i < len(proj_titles) else "",
            "tech":  proj_techs[i]  if i < len(proj_techs)  else "",
            "desc":  proj_descs[i]  if i < len(proj_descs)  else "",
        })

    cert_titles  = f.getlist("cert_title[]")
    cert_issuers = f.getlist("cert_issuer[]")
    certifications = []
    for i in range(len(cert_titles)):
        certifications.append({
            "title":  cert_titles[i]  if i < len(cert_titles)  else "",
            "issuer": cert_issuers[i] if i < len(cert_issuers) else "",
        })

    resume_data = {
        "name":           f.get("name","").strip(),
        "job_title":      f.get("job_title","").strip(),
        "email":          f.get("email","").strip(),
        "phone":          f.get("phone","").strip(),
        "location":       f.get("location","").strip(),
        "linkedin":       f.get("linkedin","").strip(),
        "summary":        f.get("summary","").strip(),
        "skills":         skills,
        "education":      education,
        "experience":     experience,
        "projects":       projects,
        "certifications": certifications,
        "languages":      f.get("languages","").strip(),
        "hobbies":        f.get("hobbies","").strip(),
    }

    if not resume_data["name"] or not resume_data["email"]:
        return render_template("builder.html",
                               error="Name and Email are required fields.")

    try:
        create_resume_pdf(resume_data, filepath="built_resume.pdf")
    except Exception as e:
        print("PDF build error:", e)
        return render_template("builder.html",
                               error=f"PDF generation failed: {str(e)}")

    # ── PDF seedha return karo (fetch() ke liye) ─────────────
    safe_name = re.sub(r'[^a-zA-Z0-9_]', '_', resume_data["name"])
    filename  = f"{safe_name}_Resume.pdf"
    return send_file("built_resume.pdf", as_attachment=True,
                     download_name=filename)





if __name__ == "__main__":
    app.run(debug=True)