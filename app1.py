from flask import Flask, render_template, request, send_file
import os
import PyPDF2
import docx
import re
from datetime import datetime
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.lib.units import mm
from reportlab.lib.enums import TA_CENTER, TA_LEFT

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# ============================================================
# 🗂️ ALL SKILLS — Domain wise (strong & complete)
# ============================================================
ALL_SKILLS = {

    # 💻 IT / Web Development
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

    # 📊 Data Science / ML / AI
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

    # 🏥 Medical / Healthcare
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

    # ⚖️ Law / Legal  ← STRONGLY EXPANDED
    "legal research": "⚖️ Law",
    "contract drafting": "⚖️ Law",
    "litigation": "⚖️ Law",
    "corporate law": "⚖️ Law",
    "criminal law": "⚖️ Law",
    "family law": "⚖️ Law",
    "intellectual property": "⚖️ Law",
    "taxation law": "⚖️ Law",
    "cyber law": "⚖️ Law",
    "constitutional law": "⚖️ Law",
    "legal writing": "⚖️ Law",
    "court proceedings": "⚖️ Law",
    "arbitration": "⚖️ Law",
    "mediation": "⚖️ Law",
    "llb": "⚖️ Law",
    "ba.llb": "⚖️ Law",
    "ballb": "⚖️ Law",
    "moot court": "⚖️ Law",
    "case analysis": "⚖️ Law",
    "case study": "⚖️ Law",
    "legal drafting": "⚖️ Law",
    "client communication": "⚖️ Law",
    "law intern": "⚖️ Law",
    "law internship": "⚖️ Law",
    "legal intern": "⚖️ Law",
    "legal internship": "⚖️ Law",
    "public speaking": "⚖️ Law",
    "ethical leadership": "⚖️ Law",
    "problem resolution": "⚖️ Law",
    "employee development": "⚖️ Law",
    "motivation techniques": "⚖️ Law",
    "self-awareness": "⚖️ Law",
    "written communication": "⚖️ Law",
    "legal compliance": "⚖️ Law",
    "advocate": "⚖️ Law",
    "paralegal": "⚖️ Law",
    "legal aid": "⚖️ Law",
    "bar council": "⚖️ Law",
    "juris": "⚖️ Law",
    "law": "⚖️ Law",

    # 📈 Finance / Accounting
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

    # 🎨 Design / Creative
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

    # 🏗️ Engineering
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

    # 📚 Education / Teaching
    "lesson planning": "📚 Education",
    "curriculum development": "📚 Education",
    "classroom management": "📚 Education",
    "student assessment": "📚 Education",
    "e-learning": "📚 Education",
    "educational technology": "📚 Education",
    "special education": "📚 Education",
    "b.ed": "📚 Education", "ctet": "📚 Education",
    "teaching": "📚 Education", "mentoring": "📚 Education",
    "tutoring": "📚 Education", "coaching": "📚 Education",

    # 📣 Marketing / HR
    "digital marketing": "📣 Marketing/HR",
    "seo": "📣 Marketing/HR",
    "social media marketing": "📣 Marketing/HR",
    "content writing": "📣 Marketing/HR",
    "email marketing": "📣 Marketing/HR",
    "google ads": "📣 Marketing/HR",
    "facebook ads": "📣 Marketing/HR",
    "brand management": "📣 Marketing/HR",
    "market research": "📣 Marketing/HR",
    "hr management": "📣 Marketing/HR",
    "recruitment": "📣 Marketing/HR",
    "payroll": "📣 Marketing/HR",
    "performance management": "📣 Marketing/HR",
    "employee relations": "📣 Marketing/HR",
    "copywriting": "📣 Marketing/HR",

    # 🔬 Science / Research
    "bioinformatics": "🔬 Science",
    "scientific writing": "🔬 Science",
    "hypothesis testing": "🔬 Science",
    "biotechnology": "🔬 Science",
    "microbiology": "🔬 Science",
    "genetics": "🔬 Science",
    "laboratory": "🔬 Science",
    "research methodology": "🔬 Science",

    # 🌐 General / Soft Skills (LOW priority — won't override domain)
    "ms office": "🌐 General",
    "microsoft office": "🌐 General",
    "microsoft excel": "🌐 General",
    "microsoft word": "🌐 General",
    "powerpoint": "🌐 General",
    "communication skills": "🌐 General",
    "leadership": "🌐 General",
    "teamwork": "🌐 General",
    "project management": "🌐 General",
    "time management": "🌐 General",
    "problem solving": "🌐 General",
    "problem-solving": "🌐 General",
    "analytical thinking": "🌐 General",
    "critical thinking": "🌐 General",
    "computer skills": "🌐 General",
}

# ============================================================
# 🗂️ JOB DATABASE
# ============================================================
JOB_DATABASE = [
    # 💻 IT
    {"job": "Frontend Developer",        "required": ["html","css","javascript","react","bootstrap"]},
    {"job": "Backend Developer",         "required": ["python","flask","django","sql","rest api"]},
    {"job": "Full Stack Developer",      "required": ["html","css","javascript","react","nodejs","mongodb","python"]},
    {"job": "Java Developer",            "required": ["java","spring boot","sql","git","rest api"]},
    {"job": "DevOps Engineer",           "required": ["docker","kubernetes","linux","aws","git"]},
    {"job": "Cloud Engineer",            "required": ["aws","azure","docker","linux","python"]},
    {"job": "Android Developer",         "required": ["java","android","kotlin","sql","git"]},
    {"job": "Software Tester / QA",      "required": ["selenium","python","sql","git"]},
    {"job": "UI/UX Designer",            "required": ["figma","adobe xd","wireframing","prototyping","canva"]},

    # 📊 Data Science
    {"job": "Data Analyst",              "required": ["python","pandas","numpy","sql","power bi"]},
    {"job": "Data Scientist",            "required": ["machine learning","python","pandas","tensorflow"]},
    {"job": "ML Engineer",               "required": ["machine learning","deep learning","tensorflow","pytorch","python"]},
    {"job": "Business Intelligence",     "required": ["power bi","tableau","sql","data analysis","microsoft excel"]},

    # 🏥 Medical
    {"job": "Medical Doctor (MBBS)",     "required": ["anatomy","physiology","pharmacology","medical diagnosis","patient care"]},
    {"job": "Nurse",                     "required": ["patient care","nursing","first aid","anatomy","physiology"]},
    {"job": "Dentist (BDS)",             "required": ["bds","anatomy","patient care","surgery","dentistry"]},
    {"job": "Pharmacist",                "required": ["pharmacology","biochemistry","patient care"]},
    {"job": "Medical Coder",             "required": ["medical coding","health informatics","anatomy"]},

    # ⚖️ Law
    {"job": "Corporate Lawyer",          "required": ["corporate law","contract drafting","legal research","legal writing"]},
    {"job": "Criminal Lawyer",           "required": ["criminal law","litigation","court proceedings","legal research"]},
    {"job": "Law Intern",                "required": ["llb","case analysis","legal research","contract drafting","client communication"]},
    {"job": "Legal Advisor",             "required": ["legal research","arbitration","mediation","legal writing"]},
    {"job": "Tax Consultant",            "required": ["taxation law","gst","income tax","legal writing"]},
    {"job": "Paralegal",                 "required": ["legal research","legal drafting","case analysis","written communication"]},

    # 📈 Finance
    {"job": "Chartered Accountant",      "required": ["accounting","auditing","gst","income tax","tally"]},
    {"job": "Financial Analyst",         "required": ["financial analysis","microsoft excel","financial modeling","investment banking"]},
    {"job": "Accountant",                "required": ["accounting","tally","gst","microsoft excel","balance sheet"]},

    # 🎨 Design
    {"job": "Graphic Designer",          "required": ["photoshop","illustrator","canva","graphic design"]},
    {"job": "Video Editor",              "required": ["premiere pro","after effects","video editing","motion graphics"]},

    # 🏗️ Engineering
    {"job": "Civil Engineer",            "required": ["autocad","staad pro","structural analysis","surveying"]},
    {"job": "Mechanical Engineer",       "required": ["solidworks","catia","ansys","matlab","mechanical design"]},

    # 📚 Education
    {"job": "School Teacher",            "required": ["teaching","lesson planning","classroom management","student assessment"]},
    {"job": "Corporate Trainer",         "required": ["communication skills","leadership","microsoft office"]},

    # 📣 Marketing / HR
    {"job": "Digital Marketing Manager", "required": ["digital marketing","seo","social media marketing","google ads","content writing"]},
    {"job": "HR Manager",                "required": ["hr management","recruitment","payroll","employee relations"]},
    {"job": "Content Writer",            "required": ["content writing","seo","communication skills"]},

    # 🔬 Science
    {"job": "Research Scientist",        "required": ["scientific writing","hypothesis testing","bioinformatics"]},
    {"job": "Biotechnologist",           "required": ["biotechnology","genetics","microbiology"]},
    {"job": "Bioinformatics Analyst",    "required": ["bioinformatics","python","data analysis","genetics"]},
    {"job": "Lab Technician",            "required": ["laboratory","microbiology","biochemistry","research methodology"]},
    {"job": "Clinical Research Assoc.",  "required": ["clinical research","scientific writing","healthcare","research methodology"]},
    {"job": "Microbiologist",            "required": ["microbiology","laboratory","genetics","biochemistry","scientific writing"]},
    {"job": "Environmental Scientist",   "required": ["research methodology","data collection","scientific writing","hypothesis testing","laboratory"]},
    {"job": "Geneticist",                "required": ["genetics","biotechnology","microbiology","laboratory","scientific writing"]},

    # 📊 Data Science — extra jobs
    {"job": "AI Engineer",               "required": ["artificial intelligence","deep learning","python","tensorflow","natural language processing"]},
    {"job": "Data Engineer",             "required": ["python","sql","big data","apache spark","hadoop"]},
    {"job": "NLP Engineer",              "required": ["natural language processing","python","deep learning","tensorflow","scikit-learn"]},
    {"job": "Computer Vision Engineer",  "required": ["computer vision","opencv","deep learning","pytorch","python"]},
    {"job": "Statistician",              "required": ["statistics","r programming","data analysis","python","data visualization"]},

    # ⚖️ Law — extra jobs
    {"job": "IP Lawyer",                 "required": ["intellectual property","legal research","contract drafting","litigation","legal writing"]},
    {"job": "Cyber Law Specialist",      "required": ["cyber law","legal research","legal writing","constitutional law"]},
    {"job": "Family Lawyer",             "required": ["family law","litigation","court proceedings","mediation","legal research"]},
    {"job": "Arbitrator",                "required": ["arbitration","mediation","legal research","legal writing","constitutional law"]},
    {"job": "Legal Researcher",          "required": ["legal research","legal writing","case analysis","written communication"]},

    # 📈 Finance — extra jobs
    {"job": "Tax Analyst",               "required": ["income tax","gst","taxation","tally","accounting"]},
    {"job": "Audit Associate",           "required": ["auditing","accounting","financial reporting","tally","ms office"]},
    {"job": "Budget Analyst",            "required": ["budgeting","financial analysis","ms excel","financial reporting","risk management"]},
    {"job": "Stock Market Analyst",      "required": ["stock market","investment banking","financial analysis","mutual funds","risk management"]},
    {"job": "Finance Manager",           "required": ["financial modeling","financial analysis","budgeting","risk management","financial reporting"]},

    # 🏥 Medical — extra jobs
    {"job": "Physiotherapist",           "required": ["anatomy","physiology","patient care","clinical skills","first aid"]},
    {"job": "Clinical Researcher",       "required": ["clinical research","pharmacology","biochemistry","patient care","medical diagnosis"]},
    {"job": "Health Informatics Spec.",  "required": ["health informatics","medical coding","patient care","anatomy"]},
    {"job": "Ayurvedic Doctor",          "required": ["ayurveda","anatomy","pharmacology","patient care","physiology"]},
    {"job": "Homeopathic Doctor",        "required": ["homeopathy","anatomy","pharmacology","patient care","physiology"]},

    # 🎨 Design — extra jobs
    {"job": "Brand Designer",            "required": ["brand design","illustrator","photoshop","canva","graphic design"]},
    {"job": "Motion Designer",           "required": ["motion graphics","after effects","premiere pro","blender"]},
    {"job": "Product Designer",          "required": ["figma","wireframing","prototyping","ui/ux","adobe xd"]},
    {"job": "Illustrator/Artist",        "required": ["illustrator","photoshop","indesign","corel draw","graphic design"]},

    # 🏗️ Engineering — extra jobs
    {"job": "Structural Engineer",       "required": ["structural analysis","autocad","staad pro","revit","surveying"]},
    {"job": "VLSI Engineer",             "required": ["vlsi","matlab","autocad","electrical engineering"]},
    {"job": "Automation Engineer",       "required": ["plc","scada","matlab","mechanical design","autocad"]},
    {"job": "CAD Designer",              "required": ["autocad","solidworks","catia","mechanical design","revit"]},

    # 📚 Education — extra jobs
    {"job": "Online Tutor",              "required": ["tutoring","e-learning","educational technology","communication skills"]},
    {"job": "Curriculum Designer",       "required": ["curriculum development","lesson planning","educational technology","student assessment"]},
    {"job": "Special Educator",          "required": ["special education","student assessment","lesson planning","mentoring"]},
    {"job": "Academic Counselor",        "required": ["mentoring","coaching","communication skills","student assessment"]},

    # 📣 Marketing/HR — extra jobs
    {"job": "Social Media Manager",      "required": ["social media marketing","content writing","digital marketing","copywriting","brand management"]},
    {"job": "Recruitment Specialist",    "required": ["recruitment","hr management","employee relations","communication skills"]},
    {"job": "Content Strategist",        "required": ["content strategy","content writing","seo","copywriting","market research"]},
    {"job": "Performance Marketer",      "required": ["google ads","facebook ads","digital marketing","market research","email marketing"]},
    {"job": "HR Business Partner",       "required": ["hr management","performance management","employee relations","payroll","recruitment"]},
]

# ============================================================
# 🌐 DOMAIN DETECTION — General skills excluded from scoring
# ============================================================
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
        # ✅ Skip General skills — they should NOT decide domain
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
# 📄 EXTRACT TEXT — Strong multi-method extraction
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
# 👤 EXTRACT BASIC INFO — Strong name, phone, email detection
# ============================================================

LOCATION_WORDS = {
    # Cities & states
    "sirsa","haryana","delhi","mumbai","bangalore","bengaluru","chennai",
    "kolkata","hyderabad","pune","jaipur","lucknow","chandigarh","amritsar",
    "ludhiana","gurugram","noida","agra","varanasi","patna","bhopal","indore",
    "nagpur","surat","ahmedabad","rajasthan","punjab","gujarat","maharashtra",
    "kerala","karnataka","uttar","pradesh","madhya","andhra","telangana",
    "bihar","jharkhand","odisha","assam","west","bengal","himachal","gurgaon",
    "faridabad","meerut","nashik","aurangabad","solapur","ranchi","raipur",
    "coimbatore","visakhapatnam","madurai","vijayawada","mysore","jodhpur",
    "india","new","north","south","east","west","central",
    # PDF/doc words
    "page","pages","resume","cv","curriculum","vitae","document","report",
    "candidate","applicant","profile","scanned","scanner","oken",
    # English common words
    "to","the","a","an","and","or","of","in","at","by","for","with",
    "is","are","was","were","be","been","has","have","had","do","does",
    "this","that","these","those","my","your","our","their","its",
    "mr","mrs","ms","dr","prof","sir","dear","from","subject","re",
    "am","i","we","they","he","she","it","who","which","what",
    # Resume section headers
    "career","objective","skills","education","experience","projects",
    "languages","summary","contact","email","phone","address","mobile",
    "references","internship","bachelor","master","computer","applications",
    "ongoing","expected","completion","developer","frontend","backend","intern",
    "about","me","declaration","hobbies","interests","achievements",
    "certifications","workshops","seminars","activities","personal","details",
    "date","birth","gender","nationality","marital","status","male","female","information","personal","curriculum","vitae2",
    "native","elementary","intermediate","fluent","proficient","english",
    "hindi","punjabi","training","activity","volunteer","nss","university",
    "student","disciplined","dedicated","skilled","motivated","passionate",
    "experienced","qualified","professional","seeking","looking","proven",
    "record","excellence","strong","ethic","highly","focused","committed",
    "consistently","delivering","quality","results","pressure","spend",
    "learn","things","fifth","year","law","intern","pursuing","currently",
    "summary","scanned","scanner","oken","powered","generated","page",
    "information","personal","curriculum","vitae",
}

def extract_basic_info(text):
    clean = text.replace("\r", " ")

    # ── EMAIL ──────────────────────────────────────────────────
    email_match = re.search(
        r"[a-zA-Z0-9][a-zA-Z0-9+_.%-]*@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}",
        clean
    )
    email = email_match.group(0) if email_match else "Not Found"

    # ── PHONE — handles +91, spaces, dashes, 10-digit ──────────
    phone_match = re.search(
        r"(\+91[\s\-]?\d{5}[\s\-]?\d{5}"      # +91 XXXXX XXXXX
        r"|\+91[\s\-]?\d{10}"                  # +91 XXXXXXXXXX
        r"|\b\d{5}[\s\-]\d{5}\b"               # XXXXX XXXXX
        r"|\b[6-9]\d{9}\b)",                   # 10-digit mobile
        clean
    )
    phone = phone_match.group(0).strip() if phone_match else "Not Found"

    # ── NAME — strong multi-strategy detection ──────────────────
    name = _detect_name(text)

    return name, phone, email


def _clean_line(line):
    """Remove noise chars from a line."""
    line = re.sub(r'[|,/\\@#*\xb7:;_=+<>{}()\[\]"\'\`]', ' ', line)
    line = re.sub(r'\s+', ' ', line).strip()
    return line

def _is_valid_name_word(w):
    """Check if word can be part of a human name."""
    return w.isalpha() and len(w) >= 2 and w.lower() not in LOCATION_WORDS

def _extract_name_from_line(line, max_words=4):
    """Try to extract a name from a single line. Returns string or None."""
    line_c = _clean_line(line)
    if not line_c:
        return None
    if re.search(r'\d', line_c):
        return None
    if len(line_c) > 50:
        return None
    words = line_c.split()
    valid = [w for w in words if _is_valid_name_word(w)]
    if not (1 <= len(valid) <= max_words):
        return None
    candidate = []
    for w in valid:
        if w[0].isupper() or w.isupper():
            candidate.append(w.title())
        else:
            break
    if not (1 <= len(candidate) <= max_words):
        return None
    joined_lower = ' '.join(candidate).lower()
    if joined_lower in LOCATION_WORDS:
        return None
    return ' '.join(candidate)


def _detect_name(text):
    """
    Ultra-strong name detection — 8 strategies.
    Fixed for:
    - Vartika: "Page 1" on first line, name on second line (ALL CAPS)
    - Raju: Name split across two lines "RAJU\nSHARMA"
    """
    lines = [l.strip() for l in text.split('\n') if l.strip()]

    # ── Pre-process: merge consecutive ALL-CAPS single-word lines ──
    # Handles "RAJU\nSHARMA" → "RAJU SHARMA"
    merged_lines = []
    i = 0
    while i < len(lines):
        line = lines[i]
        lc = _clean_line(line)
        words = lc.split()
        # If current line is single ALL-CAPS word and next line is also single ALL-CAPS word
        if (len(words) == 1 and words[0].isupper() and words[0].isalpha()
                and len(words[0]) >= 2 and words[0].lower() not in LOCATION_WORDS):
            # Check next line
            if i + 1 < len(lines):
                next_lc = _clean_line(lines[i+1])
                next_words = next_lc.split()
                if (len(next_words) == 1 and next_words[0].isupper()
                        and next_words[0].isalpha() and len(next_words[0]) >= 2
                        and next_words[0].lower() not in LOCATION_WORDS):
                    # Merge into one line
                    merged_lines.append(words[0] + ' ' + next_words[0])
                    i += 2
                    continue
        merged_lines.append(line)
        i += 1

    # Use merged lines for detection
    all_lines = merged_lines + lines  # try merged first, then original

    # ══ STRATEGY 1 ══════════════════════════════════════════════
    # First short line (1-4 alpha words, starts uppercase)
    # Skip "Page N" lines automatically
    for line in all_lines[:12]:
        # Skip "Page 1", "Page 2" etc
        if re.match(r'^page\s*\d*$', line.strip().lower()):
            continue
        name = _extract_name_from_line(line, max_words=4)
        if name:
            return name

    # ══ STRATEGY 2 ══════════════════════════════════════════════
    # ALL-CAPS line anywhere in first 15 lines (handles "VARTIKA ARORA")
    for line in lines[:15]:
        # Skip "Page N"
        if re.match(r'^page\s*\d*$', line.strip().lower()):
            continue
        lc = _clean_line(line)
        if not lc or re.search(r'\d', lc) or len(lc) > 45:
            continue
        words = lc.split()
        if (len(words) >= 1 and
                all(w.isupper() and w.isalpha() and len(w) >= 2 for w in words)):
            valid = [w.title() for w in words if w.lower() not in LOCATION_WORDS]
            if 1 <= len(valid) <= 4:
                return ' '.join(valid)

    # ══ STRATEGY 3 ══════════════════════════════════════════════
    # Line just BEFORE email address
    email_pat = re.compile(r'[a-zA-Z0-9][a-zA-Z0-9+_.%-]*@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')
    email_m = email_pat.search(text)
    if email_m:
        before  = text[:email_m.start()]
        blines  = [l.strip() for l in before.split('\n') if l.strip()]
        for line in reversed(blines[-6:]):
            if re.match(r'^page\s*\d*$', line.strip().lower()):
                continue
            name = _extract_name_from_line(line, max_words=4)
            if name:
                return name
        # Also check ALL CAPS in before lines
        for line in reversed(blines[-6:]):
            if re.match(r'^page\s*\d*$', line.strip().lower()):
                continue
            lc = _clean_line(line)
            if not lc or re.search(r'\d', lc) or len(lc) > 45:
                continue
            words = lc.split()
            if all(w.isupper() and w.isalpha() and len(w) >= 2 for w in words):
                valid = [w.title() for w in words if w.lower() not in LOCATION_WORDS]
                if 1 <= len(valid) <= 4:
                    return ' '.join(valid)

    # ══ STRATEGY 4 ══════════════════════════════════════════════
    # Line just BEFORE phone number
    phone_pat = re.compile(r'(\+91[\s\-]?\d{5}[\s\-]?\d{5}|\+91[\s\-]?\d{10}|\b[6-9]\d{9}\b)')
    phone_m = phone_pat.search(text)
    if phone_m:
        before  = text[:phone_m.start()]
        blines  = [l.strip() for l in before.split('\n') if l.strip()]
        for line in reversed(blines[-6:]):
            if re.match(r'^page\s*\d*$', line.strip().lower()):
                continue
            name = _extract_name_from_line(line, max_words=4)
            if name:
                return name
        # ALL CAPS check before phone
        for line in reversed(blines[-8:]):
            lc = _clean_line(line)
            if not lc or re.search(r'\d', lc) or len(lc) > 45:
                continue
            words = lc.split()
            if all(w.isupper() and w.isalpha() and len(w) >= 2 for w in words):
                valid = [w.title() for w in words if w.lower() not in LOCATION_WORDS]
                if 1 <= len(valid) <= 4:
                    return ' '.join(valid)

    # ══ STRATEGY 5 ══════════════════════════════════════════════
    # "Name: XYZ" label pattern
    label_pat = re.compile(
        r'(?:name|full\s*name|candidate\s*name)\s*[:\-]\s*([A-Za-z][A-Za-z\s]{1,40})',
        re.IGNORECASE
    )
    lm = label_pat.search(text)
    if lm:
        raw   = lm.group(1).strip()
        words = raw.split()
        valid = [w.title() for w in words
                 if w.isalpha() and len(w) >= 2
                 and w.lower() not in LOCATION_WORDS
                 and w.lower() not in ('name','full','candidate')]
        if valid:
            return ' '.join(valid[:3])

    # ══ STRATEGY 6 ══════════════════════════════════════════════
    # Two consecutive Title-case words anywhere in first 25 lines
    for line in all_lines[:25]:
        if re.match(r'^page\s*\d*$', line.strip().lower()):
            continue
        lc = _clean_line(line)
        if not lc or re.search(r'\d', lc) or len(lc) > 60:
            continue
        words = lc.split()
        for i in range(len(words) - 1):
            w1, w2 = words[i], words[i+1]
            if (w1.isalpha() and w2.isalpha()
                    and w1[0].isupper() and w2[0].isupper()
                    and len(w1) >= 2 and len(w2) >= 2
                    and w1.lower() not in LOCATION_WORDS
                    and w2.lower() not in LOCATION_WORDS):
                return w1.title() + ' ' + w2.title()

    # ══ STRATEGY 7 ══════════════════════════════════════════════
    # Merged consecutive ALL-CAPS single words (Raju Sharma case)
    for line in merged_lines[:15]:
        lc = _clean_line(line)
        if not lc or re.search(r'\d', lc) or len(lc) > 40:
            continue
        words = lc.split()
        if (len(words) >= 2 and
                all(w.isupper() and w.isalpha() and len(w) >= 2 for w in words)):
            valid = [w.title() for w in words if w.lower() not in LOCATION_WORDS]
            if 1 <= len(valid) <= 4:
                return ' '.join(valid)

    # ══ STRATEGY 8 ══════════════════════════════════════════════
    # Single valid Title-case word (last resort)
    for line in all_lines[:12]:
        if re.match(r'^page\s*\d*$', line.strip().lower()):
            continue
        lc = _clean_line(line)
        if not lc or re.search(r'\d', lc):
            continue
        for w in lc.split():
            if (w.isalpha() and len(w) >= 3
                    and w[0].isupper()
                    and w.lower() not in LOCATION_WORDS):
                return w.title()

    return 'Not Found'


# ============================================================
# ⚡ SKILL DETECTION — Strong with domain-aware keywords
# ============================================================

# Skills needing strict word-boundary (short/ambiguous)
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

    # Deduplicate preserving order
    seen, unique = set(), []
    for s in found:
        if s not in seen:
            seen.add(s)
            unique.append(s)

    return unique[:25]

def extract_raw_skills(text):
    return []  # kept for compatibility

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
# ── Domain → allowed job categories ────────────────────────
DOMAIN_JOB_FILTER = {
    "💻 IT":            ["Frontend Developer","Backend Developer","Full Stack Developer",
                         "Java Developer","DevOps Engineer","Cloud Engineer",
                         "Android Developer","Software Tester / QA","UI/UX Designer"],
    "📊 Data Science":  ["Data Analyst","Data Scientist","ML Engineer","AI Researcher",
                         "Business Intelligence","AI Engineer","Data Engineer",
                         "NLP Engineer","Computer Vision Engineer","Statistician"],
    "🏥 Medical":       ["Medical Doctor (MBBS)","Nurse","Dentist (BDS)","Pharmacist",
                         "Medical Coder","Radiologist","Physiotherapist",
                         "Clinical Researcher","Health Informatics Spec.",
                         "Ayurvedic Doctor","Homeopathic Doctor"],
    "⚖️ Law":           ["Corporate Lawyer","Criminal Lawyer","Law Intern","Legal Advisor",
                         "Tax Consultant","Paralegal","IP Lawyer","Cyber Law Specialist",
                         "Family Lawyer","Arbitrator","Legal Researcher"],
    "📈 Finance":       ["Chartered Accountant","Financial Analyst","Accountant",
                         "Investment Banker","Risk Analyst","Tax Analyst",
                         "Audit Associate","Budget Analyst",
                         "Stock Market Analyst","Finance Manager"],
    "🎨 Design":        ["UI/UX Designer","Graphic Designer","Video Editor","3D Artist",
                         "Brand Designer","Motion Designer","Product Designer",
                         "Illustrator/Artist"],
    "🏗️ Engineering":  ["Civil Engineer","Mechanical Engineer","BIM Engineer",
                         "Electrical Engineer","Structural Engineer",
                         "VLSI Engineer","Automation Engineer","CAD Designer"],
    "📚 Education":     ["School Teacher","Corporate Trainer","E-Learning Developer",
                         "Online Tutor","Curriculum Designer","Special Educator",
                         "Academic Counselor"],
    "📣 Marketing/HR":  ["Digital Marketing Manager","SEO Specialist","HR Manager",
                         "Content Writer","Brand Manager","Social Media Manager",
                         "Recruitment Specialist","Content Strategist",
                         "Performance Marketer","HR Business Partner"],
    "🔬 Science":       ["Research Scientist","Biotechnologist","Bioinformatics Analyst",
                         "Lab Technician","Clinical Research Assoc.",
                         "Microbiologist","Environmental Scientist","Geneticist"],
    "🌐 General":       [],   # show all if domain unknown
}

def match_jobs(skills, domain="🌐 General"):
    skills_lower = [s.lower() for s in skills]

    # Find which short domain key matches
    # Use text-based matching to avoid emoji encoding issues
    domain_lower = domain.lower()
    allowed_jobs = None

    if "law" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["⚖️ Law"]
    elif "finance" in domain_lower or "accounting" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["📈 Finance"]
    elif "medical" in domain_lower or "health" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["🏥 Medical"]
    elif "data science" in domain_lower or "ml" in domain_lower or "ai" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["📊 Data Science"]
    elif "design" in domain_lower or "creative" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["🎨 Design"]
    elif "engineering" in domain_lower or "mech" in domain_lower or "civil" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["🏗️ Engineering"]
    elif "education" in domain_lower or "teaching" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["📚 Education"]
    elif "marketing" in domain_lower or "hr" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["📣 Marketing/HR"]
    elif "science" in domain_lower or "research" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["🔬 Science"]
    elif "it" in domain_lower or "web" in domain_lower or "development" in domain_lower:
        allowed_jobs = DOMAIN_JOB_FILTER["💻 IT"]
    else:
        allowed_jobs = []  # General — show all

    results = []
    for job in JOB_DATABASE:
        # ✅ Filter: only show jobs relevant to detected domain
        if allowed_jobs is not None and len(allowed_jobs) > 0:
            if job["job"] not in allowed_jobs:
                continue

        required = job["required"]
        matched  = [r for r in required if any(r in s for s in skills_lower)]
        missing  = [r for r in required if r not in matched]
        pct      = int((len(matched) / len(required)) * 100) if required else 0
        results.append({
            "job":     job["job"],
            "match":   pct,
            "required":required,
            "missing": missing,
            "is_full": (pct == 100),
        })

    results.sort(key=lambda x: x["match"], reverse=True)
    return results

# ============================================================
# 💡 SUGGESTIONS
# ============================================================
DOMAIN_TIPS = {
    "💻 IT / Web Development": {
        "platform": "freeCodeCamp / Udemy / LeetCode",
        "tips": [
            "Build 2-3 portfolio projects and put them on GitHub",
            "Practice DSA on LeetCode or HackerRank daily",
            "Get AWS Cloud Practitioner certification",
            "Contribute to open-source projects",
            "Add your GitHub profile link on resume",
        ],
    },
    "📊 Data Science / ML / AI": {
        "platform": "Kaggle / Coursera / fast.ai",
        "tips": [
            "Participate in a Kaggle competition",
            "Build end-to-end ML projects on GitHub",
            "Get Google Data Analytics or IBM Data Science certificate",
            "Learn SQL deeply — it is essential for all data roles",
            "Create a project using real datasets (UCI, Kaggle)",
        ],
    },
    "🏥 Medical / Healthcare": {
        "platform": "NMC Guidelines / PubMed / USMLE",
        "tips": [
            "Complete clinical internship hours",
            "Learn medical coding (ICD-10, CPT codes)",
            "Stay updated with latest clinical research",
            "Add any specialization or certification to resume",
            "Mention any research papers or case studies",
        ],
    },
    "⚖️ Law / Legal": {
        "platform": "SCC Online / Manupatra / LexisNexis",
        "tips": [
            "Participate in moot court competitions",
            "Do internship at a reputed law firm",
            "Learn legal research tools (Manupatra, SCC Online)",
            "Specialize in a niche area of law (criminal, corporate, cyber)",
            "Build strong legal writing and contract drafting skills",
        ],
    },
    "📈 Finance / Accounting": {
        "platform": "ICAI / NSE Academy / Coursera",
        "tips": [
            "Pursue CA / CFA / CPA certification",
            "Master advanced Excel and financial modeling",
            "Stay updated with latest GST and income tax laws",
            "Practice with Tally ERP and SAP",
            "Add any internship or articleship experience",
        ],
    },
    "🎨 Design / Creative": {
        "platform": "Behance / Dribbble / Adobe Learn",
        "tips": [
            "Build a strong Behance or Dribbble portfolio",
            "Learn Figma for modern UI/UX design",
            "Study color theory and typography fundamentals",
            "Do freelance projects to build real experience",
            "Add your portfolio link in your resume",
        ],
    },
    "🏗️ Civil / Mech. Engineering": {
        "platform": "NPTEL / Coursera / AutoDesk Learn",
        "tips": [
            "Get certified in AutoCAD, Revit, or STAAD Pro",
            "Do internship with construction or manufacturing firm",
            "Learn BIM software for civil engineering roles",
            "Prepare for GATE if targeting PSU jobs",
        ],
    },
    "📚 Education / Teaching": {
        "platform": "DIKSHA / Coursera / NCERT Portal",
        "tips": [
            "Get B.Ed or CTET certification",
            "Learn e-learning platforms (Moodle, Google Classroom)",
            "Build engaging and structured lesson plans",
            "Take public speaking or communication courses",
        ],
    },
    "📣 Marketing / HR": {
        "platform": "Google Digital Garage / HubSpot / LinkedIn Learning",
        "tips": [
            "Get Google Analytics or HubSpot certification (free)",
            "Build your personal social media brand",
            "Learn basic SEO and content strategy",
            "Mention campaign results with numbers (e.g. 30% growth)",
        ],
    },
    "🔬 Science / Research": {
        "platform": "PubMed / ResearchGate / Coursera",
        "tips": [
            "Publish research in indexed journals",
            "Learn data analysis tools for research (SPSS, R, Python)",
            "Apply for funded research or fellowship programs",
            "Mention lab instruments and techniques you know",
        ],
    },
}

def generate_suggestions(skills, missing_skills, domain):
    lines = []
    lines.append(f"## Domain Detected: {domain}\n")
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
    lines.append("\n## General Resume Tips:")
    lines.append("- Add a strong professional summary at the top")
    lines.append("- Quantify your achievements (e.g. improved performance by 30%)")
    lines.append("- Keep resume clean and to 1-2 pages only")
    lines.append("- Add LinkedIn profile and GitHub or portfolio link")
    lines.append("- Use action verbs: Built, Designed, Managed, Led, Developed")
    lines.append("- Proofread for spelling and grammar mistakes")
    return "\n".join(lines)

# ============================================================
# 📄 BEAUTIFUL PDF REPORT
# ============================================================
def create_pdf(score, grade, skills, suggestion,
               name="N/A", phone="N/A", email="N/A", domain="General"):

    doc = SimpleDocTemplate(
        "report.pdf",
        leftMargin=15*mm, rightMargin=15*mm,
        topMargin=12*mm, bottomMargin=12*mm
    )

    style = getSampleStyleSheet()
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
    body_style  = ParagraphStyle("body", fontSize=9, fontName="Helvetica",
        textColor=DARK, leading=14)
    head_style  = ParagraphStyle("head", fontSize=10, fontName="Helvetica-Bold",
        textColor=BLUE, spaceBefore=8, spaceAfter=3)
    blt_style   = ParagraphStyle("blt", fontSize=9, fontName="Helvetica",
        textColor=colors.HexColor("#1e293b"), leading=14, leftIndent=10)

    story = []

    # Header
    story.append(Table([[Paragraph("AI Resume Analysis Report", title_style)]],
        colWidths=[180*mm]))
    story[-1].setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),DARK),
        ("TOPPADDING",(0,0),(-1,-1),14),("BOTTOMPADDING",(0,0),(-1,-1),8),
        ("LEFTPADDING",(0,0),(-1,-1),10),
    ]))
    date_str = datetime.now().strftime("%d %B %Y  |  %I:%M %p")
    story.append(Table([[Paragraph(f"Generated on: {date_str}", sub_style)]],
        colWidths=[180*mm]))
    story[-1].setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),colors.HexColor("#1e293b")),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),8),
    ]))
    story.append(Spacer(1,8))

    # Candidate Info
    story.append(Paragraph("Candidate Information", sec_style))
    story.append(HRFlowable(width="100%", thickness=1, color=CYAN, spaceAfter=6))
    info_rows = [
        [Paragraph("Name",lbl_style),  Paragraph(str(name),val_style),
         Paragraph("Phone",lbl_style), Paragraph(str(phone),val_style)],
        [Paragraph("Email",lbl_style), Paragraph(str(email),val_style),
         Paragraph("Domain",lbl_style),Paragraph(str(domain),val_style)],
    ]
    t = Table(info_rows, colWidths=[22*mm,68*mm,22*mm,68*mm])
    t.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(-1,-1),LIGHT),("GRID",(0,0),(-1,-1),0.4,SILVER),
        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7),
        ("LEFTPADDING",(0,0),(-1,-1),8),
    ]))
    story.append(t); story.append(Spacer(1,10))

    # Score & Grade
    story.append(Paragraph("Score & Grade", sec_style))
    story.append(HRFlowable(width="100%", thickness=1, color=CYAN, spaceAfter=6))
    grade_color = (GREEN if score>=8 else BLUE if score>=6
                   else colors.HexColor("#f59e0b") if score>=4
                   else colors.HexColor("#ef4444"))
    sc_lbl = ParagraphStyle("sc",fontSize=20,fontName="Helvetica-Bold",
                             textColor=WHITE,alignment=TA_CENTER)
    gr_lbl = ParagraphStyle("gr",fontSize=20,fontName="Helvetica-Bold",
                             textColor=WHITE,alignment=TA_CENTER)
    sc_sub = ParagraphStyle("scsub",fontSize=9,fontName="Helvetica",
                             textColor=colors.HexColor("#94a3b8"),alignment=TA_CENTER)
    sd = [[
        Table([[Paragraph(f"{score} / 10",sc_lbl)],[Paragraph("Resume Score",sc_sub)]],colWidths=[85*mm]),
        Table([[Paragraph(f"  {grade}  ",gr_lbl)],[Paragraph("Grade",sc_sub)]],colWidths=[85*mm]),
    ]]
    st = Table(sd, colWidths=[90*mm,90*mm])
    st.setStyle(TableStyle([
        ("BACKGROUND",(0,0),(0,0),BLUE),("BACKGROUND",(1,0),(1,0),grade_color),
        ("TOPPADDING",(0,0),(-1,-1),10),("BOTTOMPADDING",(0,0),(-1,-1),10),
        ("LEFTPADDING",(0,0),(-1,-1),6),("INNERGRID",(0,0),(-1,-1),0,WHITE),
    ]))
    story.append(st); story.append(Spacer(1,10))

    # Skills
    story.append(Paragraph("Detected Skills", sec_style))
    story.append(HRFlowable(width="100%", thickness=1, color=CYAN, spaceAfter=6))
    if skills:
        sk_style = ParagraphStyle("sk",fontSize=9,fontName="Helvetica",
                                   textColor=colors.HexColor("#0369a1"))
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
        kt.setStyle(TableStyle([
            ("BACKGROUND",(0,0),(-1,-1),LIGHT),("GRID",(0,0),(-1,-1),0.3,SILVER),
            ("TOPPADDING",(0,0),(-1,-1),6),("BOTTOMPADDING",(0,0),(-1,-1),6),
            ("LEFTPADDING",(0,0),(-1,-1),8),
        ]))
        story.append(kt)
    else:
        story.append(Paragraph("No skills detected.", body_style))
    story.append(Spacer(1,10))

    # Suggestions
    story.append(Paragraph("Suggestions & Recommendations", sec_style))
    story.append(HRFlowable(width="100%", thickness=1, color=CYAN, spaceAfter=6))
    for line in suggestion.split("\n"):
        line = line.strip()
        if not line:
            story.append(Spacer(1,3))
        elif line.startswith("##"):
            story.append(Paragraph(f"▸  {line.replace('##','').strip()}", head_style))
        elif line.startswith("-"):
            story.append(Paragraph(f"  •  {line[1:].strip()}", blt_style))
        else:
            story.append(Paragraph(line, body_style))
    story.append(Spacer(1,14))

    # Footer
    story.append(HRFlowable(width="100%", thickness=1, color=SILVER))
    ft_style = ParagraphStyle("ft",fontSize=8,fontName="Helvetica",
                               textColor=GREY,alignment=TA_CENTER,spaceBefore=6)
    story.append(Paragraph(
        "Generated by AI Resume Analyzer  •  Developed by Kirandeep Kaur  •  Confidential",
        ft_style))

    doc.build(story)
    print("PDF created ✅")

# ============================================================
# 🌐 FLASK ROUTES
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

    text = extract_text(path)
    print("TEXT PREVIEW:", text[:400])

    name, phone, email = extract_basic_info(text)
    print("BASIC INFO:", name, phone, email)

    skills = detect_skills_ai(text)
    print("SKILLS FOUND:", skills)

    score  = calculate_score(skills)
    grade  = calculate_grade(score)
    domain = detect_domain(skills)
    print("DOMAIN:", domain)

    job_results = match_jobs(skills, domain)
    # IT domain ke liye top 3 + other 6 same rehta hai
    # Baaki domains ke liye top 3 + other 6
    domain_lower = domain.lower()
    top_jobs    = job_results[:3]
    if "it" in domain_lower or "web" in domain_lower:
        other_jobs = job_results[3:9]   # IT — same as before
    else:
        other_jobs = job_results[3:9]   # All domains — top3 + 6 others

    all_missing = set()
    for j in job_results[:5]:
        for s in j["missing"]:
            all_missing.add(s)

    suggestion = generate_suggestions(skills, list(all_missing), domain)
    create_pdf(score, grade, skills, suggestion,
               name=name, phone=phone, email=email, domain=domain)

    return render_template(
        "result.html",
        top_jobs      = top_jobs,
        other_jobs    = other_jobs,
        user_skills   = skills,
        mapped_skills = {},
        score         = score,
        grade         = grade,
        suggestion    = suggestion,
        domain        = domain,
        name          = name,
        phone         = phone,
        email         = email,
    )

@app.route("/download")
def download():
    return send_file("report.pdf", as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)