from __future__ import annotations

import csv
import datetime as dt
import hashlib
import io
import json
import random
import re
from typing import Any

import streamlit as st

try:
    import fitz  # PyMuPDF
except ImportError as exc:  # pragma: no cover - dependency is listed below
    raise RuntimeError("PyMuPDF is required. Install requirements.txt first.") from exc

try:
    import docx2txt
except ImportError:  # pragma: no cover - optional until MCQs are uploaded
    docx2txt = None

try:
    from pypdf import PdfReader
except ImportError:  # pragma: no cover - dependency is listed below
    PdfReader = None


APP_TITLE = "Pediatric Hem/Onc Atlas Study Hub"
APP_VERSION = "1.1"

st.set_page_config(page_title=APP_TITLE, layout="wide")


# The topics deliberately follow an exam-first order. The page finder uses the
# English terms because the uploaded atlases are English-language books.
TOPICS: list[dict[str, Any]] = [
    {
        "id": "foundation",
        "label": "1 — الأساسيات ونضج خلايا الدم",
        "priority": "Must Know",
        "keywords": [
            "hematopoiesis", "maturation", "peripheral blood", "bone marrow",
            "blood film", "blood smear", "erythroblast", "myelocyte", "metamyelocyte",
            "neutrophil maturation", "lymphocyte maturation",
        ],
        "description": "تعرف الطبيعي أولاً: حجم الخلية، النواة، الكروماتين، السيتوبلازم، والحبيبات.",
        "pearls": [
            "وصف الصورة قبل التشخيص: RBC ثم WBC ثم platelets ثم الخلفية.",
            "حدد مرحلة النضج من شكل النواة والكروماتين ووجود nucleoli والحبيبات.",
            "لا تجعل كلمة واحدة في الصورة تلغي الوصف المورفولوجي الكامل.",
        ],
    },
    {
        "id": "rbc_morphology",
        "label": "2 — RBC morphology وRBC inclusions",
        "priority": "Must Know",
        "keywords": [
            "erythrocyte", "red cell", "anisocytosis", "poikilocytosis", "microcyte",
            "macrocyte", "hypochromia", "target cell", "spherocyte", "schistocyte",
            "elliptocyte", "teardrop", "tear drop", "rouleaux", "inclusions in erythrocytes",
            "basophilic stippling", "howell-jolly", "pappenheimer", "polychrom",
        ],
        "description": "أعلى عائد بصري: الحجم، اللون، الشكل، التوزيع، والـ inclusions.",
        "pearls": [
            "Schistocytes تعني fragmentation وتحتاج ربطها بالسياق: MAHA/DIC/TMA.",
            "Spherocytes بلا central pallor؛ فكر في HS أو immune hemolysis ثم استخدم DAT والسياق.",
            "Target cells ليست تشخيصاً وحدها؛ اربطها بالهيموغلوبينوباثي أو الكبد أو نقص الطحال.",
        ],
    },
    {
        "id": "anemia",
        "label": "3 — Anemia وhemoglobin disorders",
        "priority": "Must Know",
        "keywords": [
            "hypochromic anemia", "iron deficiency", "thalassemia", "sickle", "hemoglobin",
            "hemoglobinopathy", "hemolytic anemia", "megaloblastic", "aplastic anemia",
            "dyserythropoietic", "porphyria", "iron overload", "glucose-6-phosphate",
            "g6pd", "hereditary spherocytosis", "autoimmune hemolytic",
        ],
        "description": "حوّل المظهر إلى pattern: microcytic، hemolytic، macrocytic، أو marrow failure.",
        "pearls": [
            "Microcytosis لا تساوي iron deficiency؛ قارن RDW، RBC count، ferritin، والـ smear.",
            "في hemolysis اجمع morphology مع reticulocytes وbilirubin وLDH وDAT.",
            "في sickle/thalassemia لا تعتمد على smear فقط؛ اربطها بالـ electrophoresis/genotype.",
        ],
    },
    {
        "id": "acute_leukemia",
        "label": "4 — Acute leukemia: ALL / AML / APL",
        "priority": "Must Know",
        "keywords": [
            "acute leukemia", "acute myeloid", "aml", "acute lymphoblastic", "all",
            "lymphoblast", "myeloblast", "blast", "promyelocytic", "auer rod", "auer rods",
            "peroxidase", "myeloperoxidase", "monoblast", "megakaryoblast", "leukemia",
        ],
        "description": "التمييز بين blast families، Auer rods، وclinical emergency patterns.",
        "pearls": [
            "Auer rods تدعم myeloid differentiation، وbundles في abnormal promyelocytes ترفع الشك بـ APL.",
            "MPO وflow cytometry وgenetics تكمل morphology؛ لا تشخّص lineage من الصورة وحدها.",
            "الـ classification الحالي لا يساوي تسميات الأطالس القديمة؛ سجّل دائماً مصدر وedition الصفحة.",
        ],
    },
    {
        "id": "chronic_myeloid",
        "label": "5 — MDS / MPN / chronic myeloid disorders",
        "priority": "High Yield",
        "keywords": [
            "myeloproliferative", "mpn", "myelodysplastic", "mds", "chronic myeloid",
            "cml", "polycythemia", "thrombocythemia", "myelofibrosis", "monocyt", "dysplasia",
        ],
        "description": "اقرأ left shift، basophilia، dysplasia، tear drops، والـ megakaryocyte patterns.",
        "pearls": [
            "وجود left shift وحده لا يساوي CML؛ اربطه بالـ basophilia والـ clinical/molecular context.",
            "Dysplasia تحتاج pattern متكرر في lineage أو أكثر، وليس خلية شاذة منفردة.",
            "استعمل WHO/ICC الحاليين عند مراجعة الأسماء والتصنيفات.",
        ],
    },
    {
        "id": "platelet_coagulation",
        "label": "6 — Platelets وbleeding disorders",
        "priority": "High Yield",
        "keywords": [
            "platelet", "thrombocytopenia", "thrombocytosis", "bleeding", "coagulation",
            "vascular", "von willebrand", "purpura", "thrombosis", "megakaryocyte",
        ],
        "description": "حجم الصفائح، giant platelets، وقراءة الصورة ضمن قصة النزف أو الخثار.",
        "pearls": [
            "Platelet size and distribution can separate production problems from peripheral destruction, but are not standalone tests.",
            "اربط smear مع platelet count، PT/aPTT، fibrinogen، وclinical phenotype.",
            "لا تستخدم morphology وحدها لنفي DIC أو TMA.",
        ],
    },
    {
        "id": "histiocytic_infection",
        "label": "7 — Histiocytic disorders وparasites/infection",
        "priority": "High Yield",
        "keywords": [
            "histiocytic", "histiocyte", "langerhans", "lch", "hemophagocytosis", "parasite",
            "malaria", "babesia", "leishmania", "microorganism", "infection", "inclusion body",
        ],
        "description": "صور لا تنساها: Langerhans، hemophagocytosis، والطفيليات داخل الدم.",
        "pearls": [
            "Hemophagocytosis قد تكون داعمة لكنها ليست وحدها كافية لتشخيص HLH.",
            "في LCH، morphology تُربط بالـ immunophenotype (CD1a/langerin) وبالسياق السريري.",
            "الطفيلي يحتاج وصف الشكل ومكانه ودورة الحياة/السياق الوبائي؛ لا تعتمد على تشابه بصري فقط.",
        ],
    },
    {
        "id": "other",
        "label": "8 — مواضيع أخرى: fluids / transfusion / miscellaneous",
        "priority": "Second Pass",
        "keywords": [
            "body fluids", "transfusion", "miscellaneous", "newborn", "nonhematopoietic",
            "tissue typing", "stem cell transplantation", "transplantation",
        ],
        "description": "تُراجع بعد تثبيت الأساسيات واللوكيميا والأنيميا.",
        "pearls": [
            "ضع هذه الصفحات في second pass إذا كان وقتك محدوداً.",
            "دوّن أي نقطة متكررة في بنك أسئلتك حتى لو كانت الصورة نفسها قليلة الورود.",
        ],
    },
]

TOPIC_BY_ID = {topic["id"]: topic for topic in TOPICS}

STUDY_ORDER = [
    ("Must Know", "Foundation + RBC morphology", "Rodak 5th: chapters 1–14"),
    ("Must Know", "Anemia + hemoglobin disorders", "Rodak 5th: chapters 10–13; Hoffbrand: chapters 5–9"),
    ("Must Know", "ALL / AML / APL", "Rodak 5th: chapters 14–16; Hoffbrand: chapter 12"),
    ("High Yield", "MDS / MPN + platelets", "Rodak 5th: chapters 17–19; Hoffbrand: chapters 13–15, 24–26"),
    ("High Yield", "Histiocytic disorders + parasites", "Rodak 5th: chapters 21–22; Hoffbrand: chapters 16, 28"),
    ("Second Pass", "Newborn, fluids, transfusion, transplant", "Hoffbrand: chapters 23, 27, 29; Rodak: chapters 23–24"),
]


def today_iso() -> str:
    return dt.date.today().isoformat()


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def init_state() -> None:
    defaults: dict[str, Any] = {
        "atlas_marks": set(),
        "atlas_red_zone": set(),
        "atlas_seen": set(),
        "atlas_quiz_results": {},
        "atlas_quiz_page": None,
        "atlas_quiz_signature": None,
        "atlas_quiz_revealed": False,
        "mcq_key": None,
        "mcq_idx": 0,
        "mcq_results": {},
        "mcq_marks": set(),
        "mcq_revealed": False,
        "mcq_review_mode": False,
        "mcq_review_indices": [],
        "restore_notice": "",
    }
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value.copy() if isinstance(value, (dict, set, list)) else value


init_state()

font_size = st.sidebar.slider("🔠 حجم الخط", min_value=15, max_value=30, value=20)
st.markdown(
    f"""
    <style>
      .block-container {{ padding-top: 1rem; padding-bottom: 1rem; }}
      footer {{ visibility: hidden; }}
      .stMarkdown p, .stRadio label, .stCheckbox label, .stTextInput label,
      .stTextArea label, .stSelectbox label {{ font-size: {font_size}px !important; line-height: 1.55 !important; }}
      button[data-baseweb="tab"] {{ font-size: 18px !important; font-weight: 700; }}
    </style>
    """,
    unsafe_allow_html=True,
)


def source_kind(filename: str) -> str:
    lowered = filename.lower()
    if "color atlas" in lowered or "hoffbrand" in lowered:
        return "hoffbrand4"
    if "رهيب" in filename or "2013" in lowered or "fourth edition" in lowered:
        return "rodak4"
    if "dr maha" in lowered or "book" in lowered or "clinical hematology atlas" in lowered:
        return "rodak5"
    return "uploaded"


def source_label(kind: str, filename: str) -> str:
    if kind == "rodak5":
        return "Rodak Clinical Hematology Atlas — Fifth Edition (PRIMARY)"
    if kind == "hoffbrand4":
        return "Hoffbrand Color Atlas — Fourth Edition (SUPPLEMENT)"
    if kind == "rodak4":
        return "Rodak Clinical Hematology Atlas — Fourth Edition (OLDER / DUPLICATE)"
    return f"Atlas uploaded by user — {filename}"


def make_source(uploaded: Any) -> dict[str, Any]:
    data = uploaded.getvalue()
    digest = hashlib.sha256(data).hexdigest()[:20]
    kind = source_kind(str(uploaded.name))
    return {
        "id": digest,
        "filename": str(uploaded.name),
        "bytes": data,
        "kind": kind,
        "label": source_label(kind, str(uploaded.name)),
    }


def clean_text(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def infer_topic(chapter: str, text: str) -> str:
    haystack = f"{chapter} {text}".lower()

    # Chapter titles are much more reliable than repeated running text. This
    # keeps a page in an anemia/leukemia chapter from being mislabelled as
    # "foundation" just because the page also mentions bone marrow.
    chapter_rules = [
        ("acute_leukemia", [
            "acute leukemia", "acute myeloid", "precursor lymphoid", "acute leukemias",
        ]),
        ("histiocytic_infection", [
            "histiocytic", "microorganism", "parasitic disorders", "parasite",
        ]),
        ("platelet_coagulation", [
            "platelet", "bleeding disorders", "coagulation", "thrombosis",
        ]),
        ("chronic_myeloid", [
            "myeloproliferative", "myelodysplastic", "chronic myeloid", "chronic lymphoid",
            "mature lymphoproliferative", "non-hodgkin", "hodgkin lymphoma", "myeloma",
        ]),
        ("anemia", [
            "hypochromic anemia", "hemolytic anemia", "megaloblastic", "aplastic",
            "dyserythropoietic", "porphyria", "iron overload", "genetic disorders of hemoglobin",
            "diseases affecting erythrocytes",
        ]),
        ("rbc_morphology", [
            "variations in size and color of erythrocytes",
            "variations in shape and distribution of erythrocytes",
            "inclusions in erythrocytes", "erythrocyte maturation",
        ]),
        ("foundation", [
            "hematopoiesis", "cellular basis of hematopoiesis", "cell machinery",
            "growth factors", "maturation of blood cells", "peripheral blood film",
            "blood film examination", "examination in peripheral blood and bone marrow",
        ]),
    ]
    chapter_lower = chapter.lower()
    for topic_id, phrases in chapter_rules:
        if any(phrase in chapter_lower for phrase in phrases):
            return topic_id

    best_id = "other"
    best_score = 0
    for topic in TOPICS:
        if topic["id"] == "other":
            continue
        score = 0
        for keyword in topic["keywords"]:
            if keyword.lower() in haystack:
                score += 1 + haystack.count(keyword.lower()) // 12
        if chapter and any(keyword.lower() in chapter.lower() for keyword in topic["keywords"]):
            score += 3
        if score > best_score:
            best_id, best_score = topic["id"], score
    return best_id


def chapter_for_page(chapters: list[dict[str, Any]], page_number: int) -> str:
    current = "Unindexed section"
    for chapter in chapters:
        if chapter["pdf_page"] <= page_number:
            current = chapter["title"]
        else:
            break
    return current


@st.cache_data(show_spinner=False, max_entries=12)
def build_pdf_index(file_bytes: bytes, filename: str) -> dict[str, Any]:
    """Build a lightweight text/bookmark index; page images are rendered on demand."""
    document = fitz.open(stream=file_bytes, filetype="pdf")
    try:
        raw_toc = document.get_toc(simple=True) or []
    except Exception:
        raw_toc = []

    chapter_entries: list[dict[str, Any]] = []
    for item in raw_toc:
        if len(item) < 3:
            continue
        level, title, page_number = item[0], clean_text(str(item[1])), safe_int(item[2], 1)
        # Level 1 is normally the chapter level. If a PDF lacks level 1
        # bookmarks, the first available level still gives useful navigation.
        if level == 1 and page_number >= 1:
            chapter_entries.append({"title": title, "pdf_page": page_number})
    if not chapter_entries:
        for item in raw_toc:
            if len(item) >= 3 and safe_int(item[2], 0) >= 1:
                chapter_entries.append({"title": clean_text(str(item[1])), "pdf_page": safe_int(item[2], 1)})
                if len(chapter_entries) >= 80:
                    break
    chapter_entries.sort(key=lambda row: row["pdf_page"])

    pages: list[dict[str, Any]] = []
    for page_index in range(len(document)):
        page = document.load_page(page_index)
        page_number = page_index + 1
        text = clean_text(page.get_text("text"))
        chapter = chapter_for_page(chapter_entries, page_number)
        topic_id = infer_topic(chapter, text)
        is_cover = bool(re.search(r"\bchapter\s+\d+\b", text.lower())) and len(text) < 220
        pages.append(
            {
                "pdf_page": page_number,
                "chapter": chapter,
                "topic_id": topic_id,
                "text": text[:8000],
                "has_text": bool(text),
                "has_images": bool(page.get_images(full=True)),
                "is_cover": is_cover,
            }
        )

    return {
        "page_count": len(document),
        "chapters": chapter_entries,
        "pages": pages,
    }


@st.cache_data(show_spinner=False, max_entries=80)
def render_pdf_page(file_bytes: bytes, pdf_page: int, crop_mode: str, zoom: float) -> bytes:
    """Render a page; crop_mode tries to isolate the main figure for self-testing."""
    document = fitz.open(stream=file_bytes, filetype="pdf")
    page = document.load_page(max(0, pdf_page - 1))
    clip = page.rect

    if crop_mode == "figure":
        try:
            blocks = page.get_text("dict").get("blocks", [])
            image_blocks = []
            for block in blocks:
                if block.get("type") != 1 or not block.get("bbox"):
                    continue
                x0, y0, x1, y1 = block["bbox"]
                area = max(0.0, x1 - x0) * max(0.0, y1 - y0)
                image_blocks.append((area, fitz.Rect(x0, y0, x1, y1)))
            if image_blocks:
                area, largest = max(image_blocks, key=lambda pair: pair[0])
                page_area = page.rect.width * page.rect.height
                if area < page_area * 0.86:
                    # Captions in Rodak are separate text blocks immediately
                    # below the figure. A vertical pad accidentally includes
                    # lines such as "FIGURE 24-20 ...", which gives away the
                    # answer during the exam challenge. Keep the exact image
                    # rectangle vertically; use only a tiny horizontal pad.
                    pad_x = largest.width * 0.02
                    pad_y = 0
                    clip = fitz.Rect(
                        max(page.rect.x0, largest.x0 - pad_x),
                        max(page.rect.y0, largest.y0 - pad_y),
                        min(page.rect.x1, largest.x1 + pad_x),
                        min(page.rect.y1, largest.y1 + pad_y),
                    )
            else:
                clip = fitz.Rect(
                    page.rect.x0,
                    page.rect.y0 + page.rect.height * 0.10,
                    page.rect.x1,
                    page.rect.y1 - page.rect.height * 0.12,
                )
        except Exception:
            clip = page.rect

    matrix = fitz.Matrix(float(zoom), float(zoom))
    pixmap = page.get_pixmap(matrix=matrix, clip=clip, alpha=False)
    return pixmap.tobytes("png")


def page_id(source: dict[str, Any], pdf_page: int) -> str:
    return f"{source['id']}:{pdf_page}"


def topic_label(topic_id: str) -> str:
    return TOPIC_BY_ID.get(topic_id, TOPIC_BY_ID["other"])["label"]


def filtered_sources(sources: list[dict[str, Any]], include_older: bool) -> list[dict[str, Any]]:
    if include_older:
        return sources
    primary = [source for source in sources if source["kind"] != "rodak4"]
    return primary or sources


def page_matches(page: dict[str, Any], topic_id: str, search: str) -> bool:
    if topic_id != "all" and page["topic_id"] != topic_id:
        return False
    if search:
        query = search.lower().strip()
        searchable = f"{page['chapter']} {page['text']}".lower()
        return query in searchable
    return True


def select_page_record(sources: list[dict[str, Any]], source_id: str, pdf_page: int) -> tuple[dict[str, Any], dict[str, Any]] | None:
    for source in sources:
        if source["id"] != source_id:
            continue
        index = build_pdf_index(source["bytes"], source["filename"])
        for page in index["pages"]:
            if page["pdf_page"] == pdf_page:
                return source, page
    return None


def atlas_progress_csv() -> bytes:
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Page ID", "Status", "Marked", "Red zone"])
    result_map = st.session_state.get("atlas_quiz_results", {})
    for key in sorted(set(st.session_state.get("atlas_seen", set())) | set(st.session_state.get("atlas_marks", set()))):
        writer.writerow([
            key,
            result_map.get(key, "seen"),
            "Yes" if key in st.session_state.get("atlas_marks", set()) else "No",
            "Yes" if key in st.session_state.get("atlas_red_zone", set()) else "No",
        ])
    return output.getvalue().encode("utf-8-sig")


def progress_payload() -> dict[str, Any]:
    return {
        "app": APP_TITLE,
        "version": APP_VERSION,
        "saved_at": dt.datetime.now().isoformat(timespec="seconds"),
        "atlas_marks": sorted(st.session_state.get("atlas_marks", set())),
        "atlas_red_zone": sorted(st.session_state.get("atlas_red_zone", set())),
        "atlas_seen": sorted(st.session_state.get("atlas_seen", set())),
        "atlas_quiz_results": dict(st.session_state.get("atlas_quiz_results", {})),
        "mcq_results": dict(st.session_state.get("mcq_results", {})),
        "mcq_marks": sorted(st.session_state.get("mcq_marks", set())),
    }


def restore_progress(raw: bytes) -> str:
    try:
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, dict):
            return "الملف ليس بصيغة تقدم صحيحة."
        for key in ("atlas_marks", "atlas_red_zone", "atlas_seen", "mcq_marks"):
            values = data.get(key, [])
            if isinstance(values, list):
                st.session_state[key] = set(str(value) for value in values)
        for key in ("atlas_quiz_results", "mcq_results"):
            values = data.get(key, {})
            if isinstance(values, dict):
                st.session_state[key] = {str(k): str(v) for k, v in values.items()}
        return "تم استرجاع التقدم بنجاح."
    except (UnicodeDecodeError, json.JSONDecodeError, TypeError):
        return "تعذر قراءة ملف التقدم."


# ------------------------- MCQ support -------------------------
QUESTION_START_RE = re.compile(r"(?im)^[ \t]*(?:(?:question|q)\s*#?\s*\d+|neufeld\s*#?\s*\d+|\d+)\s*[:.)\-]?\s*")
ANSWER_HEADING_RE = re.compile(r"(?im)^[ \t]*(?:✅\s*)?(?:the\s+)?correct\s+answer\s*(?:is\s*)?(?:[:.\-]\s*)?")
EXPLANATION_RE = re.compile(r"(?im)^[ \t]*(?:💡\s*)?(?:explanation|discussion|rationale|clinical\s+rationale)\s*[:.\-]?\s*")


def extract_mcq_text(data: bytes, filename: str) -> str:
    if filename.lower().endswith(".docx"):
        if docx2txt is None:
            return ""
        return docx2txt.process(io.BytesIO(data)) or ""
    if PdfReader is None:
        return ""
    reader = PdfReader(io.BytesIO(data))
    return "\n".join((page.extract_text() or "") for page in reader.pages)


def answer_letter(text: str) -> str:
    patterns = [
        r"(?i)\b(?:the\s+)?correct\s+answer\s*(?:is|:|-)\s*\(?([A-E])\)?\b",
        r"(?i)\banswer\s*(?:is|:|-)\s*\(?([A-E])\)?\b",
        r"(?i)\bcorrect\s+(?:option|choice)\s*(?:is|:|-)\s*\(?([A-E])\)?\b",
        r"(?i)^\s*[\[(]?([A-E])[\])]?(?=\s*(?:[.:)\-]|$))",
    ]
    for pattern in patterns:
        match = re.search(pattern, text)
        if match:
            return match.group(1).upper()
    return "?"


def parse_mcq_file(data: bytes, filename: str) -> list[dict[str, Any]]:
    text = extract_mcq_text(data, filename)
    if not text.strip():
        return []
    questions: list[dict[str, Any]] = []
    markers = list(ANSWER_HEADING_RE.finditer(text))
    if markers:
        for marker_index, marker in enumerate(markers):
            previous_end = markers[marker_index - 1].end() if marker_index else 0
            before = text[previous_end:marker.start()]
            starts = list(QUESTION_START_RE.finditer(before))
            q_text = before[starts[-1].start():].strip() if starts else before.strip()
            next_start = QUESTION_START_RE.search(text, marker.end())
            end = next_start.start() if next_start else len(text)
            answer_blob = text[marker.end():end].strip()
            split = EXPLANATION_RE.search(answer_blob)
            explanation = answer_blob[split.end():].strip() if split else ""
            answer_part = answer_blob[:split.start()].strip() if split else answer_blob
            if len(q_text) >= 15:
                questions.append({"q": q_text, "answer": answer_part, "expl": explanation, "letter": answer_letter(answer_part), "source": filename})
    if questions:
        return questions

    starts = list(QUESTION_START_RE.finditer(text))
    for index, start in enumerate(starts):
        end = starts[index + 1].start() if index + 1 < len(starts) else len(text)
        chunk = text[start.end():end].strip()
        split = EXPLANATION_RE.search(chunk)
        explanation = chunk[split.end():].strip() if split else ""
        question_part = chunk[:split.start()].strip() if split else chunk
        answer_match = re.search(r"(?i)\b(?:correct\s+answer|answer)\s*(?:is|:|-)\s*\(?([A-E])\)?\b", question_part)
        letter = answer_match.group(1).upper() if answer_match else answer_letter(explanation)
        if answer_match:
            question_part = question_part[:answer_match.start()].strip()
        if len(question_part) >= 15:
            questions.append({"q": question_part, "answer": f"Correct answer: {letter}" if letter != "?" else "Answer not detected", "expl": explanation, "letter": letter, "source": filename})
    return questions


def mcq_status(index: int) -> str | None:
    value = st.session_state.get("mcq_results", {}).get(str(index))
    return str(value) if value else None


def reset_mcq_state(key: str) -> None:
    st.session_state.mcq_key = key
    st.session_state.mcq_idx = 0
    st.session_state.mcq_results = {}
    st.session_state.mcq_marks = set()
    st.session_state.mcq_revealed = False
    st.session_state.mcq_review_mode = False
    st.session_state.mcq_review_indices = []


def mcq_review_indices(total: int, mode: str) -> list[int]:
    wrong = [i for i in range(total) if mcq_status(i) == "wrong"]
    marked = sorted(i for i in st.session_state.get("mcq_marks", set()) if 0 <= i < total)
    if mode == "wrong":
        return wrong
    if mode == "marked":
        return marked
    result: list[int] = []
    for index in wrong + marked:
        if index not in result:
            result.append(index)
    return result


def render_plan(sources: list[dict[str, Any]]) -> None:
    st.header("🧭 خطة دراسة الأطلس للبورد")
    st.write("الهدف من هذه النسخة أن تقرأ الصورة كحالة امتحانية، مو مجرد تقليب صفحات.")

    if sources:
        st.success(f"تم تحميل {len(sources)} ملف/ملفات أطلس داخل هذه الجلسة.")
        for source in sources:
            st.caption(source["label"])
    else:
        st.info("ابدأ برفع ملفي الأطلس من الشريط الجانبي. النسخة لا تحتاج رفع نسخة Rodak الرابعة عادةً.")

    st.subheader("ترتيب الأولويات")
    st.table([
        {"الأولوية": priority, "المحور": focus, "أماكن البداية": pages}
        for priority, focus, pages in STUDY_ORDER
    ])

    st.subheader("طريقة كل صورة — 60 ثانية")
    st.markdown(
        """
1. **وصف بلا تشخيص:** حجم وشكل ولون RBC، ثم WBC، ثم platelets.
2. **التقط العلامة المفتاحية:** Auer rod؟ schistocyte؟ spherocyte؟ blast؟ parasite؟
3. **اذكر احتمالين أو ثلاثة:** لا تقفز إلى تشخيص واحد قبل السياق.
4. **اربطها بالاختبار:** CBC/reticulocyte/DAT/flow/molecular حسب السؤال.
5. **سجّل سبب الخطأ:** معلومة ناقصة، تشابه صورتين، أو قراءة سريعة.
        """
    )

    st.subheader("الخطة العملية اليومية")
    st.markdown(
        """
- جلسة أولى: **20–30 صورة** من محور واحد، كل صورة مع وصف مكتوب.
- جلسة ثانية: **10–20 سؤال MCQ** من بنكك.
- جلسة أخيرة: مراجعة الصور التي وضعتها في **القبو الأحمر** فقط.
- لا تخلط Rodak 4th و5th كأنهما مصدران مستقلان؛ استعمل الخامس كأساس والثاني للمقارنة عند الحاجة.
        """
    )

    st.warning(
        "تنبيه تصنيفي: الأطلس مفيد للمورفولوجي، لكنه لا يكفي وحده للتصنيف الحديث أو العلاج. "
        "راجع WHO Classification of Tumours Online وWHO 5th/ICC عند نقاط اللوكيميا والتصنيف."
    )
    st.markdown(
        "[WHO Classification of Tumours Online](https://tumourclassification.iarc.who.int/) · "
        "[WHO Haematolymphoid Tumours, 5th edition](https://publications.iarc.who.int/Book-And-Report-Series/Who-Classification-Of-Tumours/Haematolymphoid-Tumours-2024) · "
        "[ICSH peripheral-blood morphology standardization](https://www.icsh.org/recommendations-for-standardization-of-nomenclature-and-grading-of-peripheral-blood-cell)"
    )


def render_atlas_explorer(sources: list[dict[str, Any]]) -> None:
    st.header("🔬 مستكشف الأطلس")
    if not sources:
        st.info("ارفع أطلساً واحداً على الأقل من الشريط الجانبي.")
        return

    include_older = st.checkbox("إظهار Rodak 4th القديم أيضاً", value=False, key="include_older_atlas")
    usable_sources = filtered_sources(sources, include_older)
    source_ids = [source["id"] for source in usable_sources]
    default_source = st.session_state.get("atlas_explorer_source")
    if default_source not in source_ids:
        default_source = source_ids[0]
    source_id = st.selectbox(
        "اختر المصدر",
        source_ids,
        index=source_ids.index(default_source),
        format_func=lambda sid: next(source["label"] for source in usable_sources if source["id"] == sid),
        key="atlas_explorer_source",
    )
    source = next(source for source in usable_sources if source["id"] == source_id)
    index = build_pdf_index(source["bytes"], source["filename"])

    topic_options = ["all"] + [topic["id"] for topic in TOPICS]
    selected_topic = st.selectbox(
        "المحور",
        topic_options,
        format_func=lambda topic_id: "كل الصفحات" if topic_id == "all" else topic_label(topic_id),
        key="atlas_explorer_topic",
    )
    search = st.text_input("🔎 بحث داخل النص المستخرج (English keyword أفضل)", key="atlas_explorer_search")
    matching_pages = [page for page in index["pages"] if page_matches(page, selected_topic, search)]
    st.caption(f"النتيجة: {len(matching_pages)} صفحة من أصل {index['page_count']} — رقم الصفحة هنا هو PDF page.")
    if not matching_pages:
        st.warning("ما لقيت صفحات بهذا الفلتر؛ جرّب كل الصفحات أو كلمة أبسط.")
        return

    page_options = [page["pdf_page"] for page in matching_pages]
    selected_page = st.selectbox(
        "اختر الصفحة",
        page_options,
        format_func=lambda number: next(
            f"PDF page {number} · {next(page['chapter'] for page in matching_pages if page['pdf_page'] == number)[:75]}"
            for page in matching_pages if page["pdf_page"] == number
        ),
        key=f"atlas_page_{source_id}_{selected_topic}_{hash(search)}",
    )
    page = next(page for page in matching_pages if page["pdf_page"] == selected_page)
    pid = page_id(source, selected_page)

    left, right = st.columns([3, 1])
    with left:
        st.markdown(f"### {topic_label(page['topic_id'])}")
        st.caption(f"{source['label']} · PDF page {selected_page} · {page['chapter']}")
    with right:
        zoom = st.slider("التكبير", 1.0, 2.2, 1.35, 0.05, key=f"atlas_zoom_{pid}")

    image = render_pdf_page(source["bytes"], selected_page, "full", zoom)
    st.image(image, use_container_width=True)
    mark_col, red_col, seen_col = st.columns(3)
    with mark_col:
        marked = st.checkbox("📌 تأشير الصفحة", value=pid in st.session_state.atlas_marks, key=f"mark_page_{pid}")
        if marked:
            st.session_state.atlas_marks.add(pid)
        else:
            st.session_state.atlas_marks.discard(pid)
    with red_col:
        red = st.checkbox("🔥 القبو الأحمر", value=pid in st.session_state.atlas_red_zone, key=f"red_page_{pid}")
        if red:
            st.session_state.atlas_red_zone.add(pid)
        else:
            st.session_state.atlas_red_zone.discard(pid)
    with seen_col:
        if st.button("✅ سجلتها كمقروءة", key=f"seen_page_{pid}", use_container_width=True):
            st.session_state.atlas_seen.add(pid)
            st.toast("تم تسجيل الصفحة.")

    with st.expander("📄 النص/العنوان المستخرج من الصفحة"):
        if page["text"]:
            st.text(page["text"])
        else:
            st.info("هذه الصفحة غالباً صورة ممسوحة أو نصها غير قابل للاستخراج؛ اعتمد على العرض البصري.")

    topic = TOPIC_BY_ID.get(page["topic_id"], TOPIC_BY_ID["other"])
    with st.expander("🩺 نقاط البورد لهذا المحور", expanded=True):
        for pearl in topic["pearls"]:
            st.markdown(f"- {pearl}")
    st.info("نصيحة: قبل قراءة العنوان، اكتب وصفك في ورقة أو في ملاحظاتك. هذا يحول المشاهدة إلى تدريب امتحاني.")


def choose_quiz_page(sources: list[dict[str, Any]], topic_id: str, source_id: str | None) -> dict[str, Any] | None:
    candidates: list[dict[str, Any]] = []
    selected_sources = [source for source in sources if source_id is None or source["id"] == source_id]
    for source in selected_sources:
        index = build_pdf_index(source["bytes"], source["filename"])
        for page in index["pages"]:
            if topic_id != "all" and page["topic_id"] != topic_id:
                continue
            if page["is_cover"] or (not page["has_text"] and not page["has_images"]):
                continue
            if len(page["text"]) < 35 and not page["has_images"]:
                continue
            candidates.append({"source_id": source["id"], "pdf_page": page["pdf_page"]})
    if not candidates:
        return None
    old = st.session_state.get("atlas_quiz_page")
    choices = [item for item in candidates if item != old] or candidates
    return random.choice(choices)


def render_image_quiz(sources: list[dict[str, Any]]) -> None:
    st.header("🧠 اختبار الصورة — شوف، وصف، ثم اكشف")
    if not sources:
        st.info("ارفع أطلساً من الشريط الجانبي حتى يبدأ اختبار الصور.")
        return

    usable_sources = filtered_sources(sources, st.session_state.get("include_older_atlas", False))
    source_options = ["all"] + [source["id"] for source in usable_sources]
    quiz_source = st.selectbox(
        "المصدر",
        source_options,
        format_func=lambda sid: "المصدران الأساسيان" if sid == "all" else next(source["label"] for source in usable_sources if source["id"] == sid),
        key="atlas_quiz_source",
    )
    quiz_topic = st.selectbox(
        "المحور",
        ["all"] + [topic["id"] for topic in TOPICS if topic["id"] != "other"],
        format_func=lambda tid: "اختيار عشوائي من الكل" if tid == "all" else topic_label(tid),
        key="atlas_quiz_topic",
    )
    signature = f"{quiz_source}:{quiz_topic}"
    if st.session_state.get("atlas_quiz_signature") != signature:
        st.session_state.atlas_quiz_signature = signature
        st.session_state.atlas_quiz_page = choose_quiz_page(usable_sources, quiz_topic, None if quiz_source == "all" else quiz_source)
        st.session_state.atlas_quiz_revealed = False

    new_col, crop_col = st.columns([1, 2])
    with new_col:
        if st.button("🎲 صورة جديدة", type="primary", key="new_atlas_quiz"):
            st.session_state.atlas_quiz_page = choose_quiz_page(usable_sources, quiz_topic, None if quiz_source == "all" else quiz_source)
            st.session_state.atlas_quiz_revealed = False
            st.rerun()
    with crop_col:
        crop = st.checkbox("إظهار منطقة الشكل فقط (قد لا يعزل كل العناوين)", value=True, key="quiz_crop")

    selected = st.session_state.get("atlas_quiz_page")
    if not selected:
        st.warning("لم أجد صفحات مناسبة لهذا الفلتر.")
        return
    record = select_page_record(usable_sources, selected["source_id"], selected["pdf_page"])
    if not record:
        st.warning("تعذر فتح الصفحة المختارة؛ اختر صورة جديدة.")
        return
    source, page = record
    pid = page_id(source, page["pdf_page"])
    st.caption(f"المصدر والصفحة مخفيان إلى حدّ كبير أثناء السؤال — PID: {pid[:10]}…")
    image = render_pdf_page(source["bytes"], page["pdf_page"], "figure" if crop else "full", 1.45)
    st.image(image, use_container_width=True)
    st.markdown("### قبل الكشف")
    st.write("اكتب 2–3 ملاحظات مورفولوجية، ثم احتمالَك الأقوى. لا تبحث عن العنوان أولاً.")
    st.text_area("وصفك", key=f"quiz_guess_{pid}", height=110)

    reveal_col, know_col, unsure_col, wrong_col = st.columns(4)
    with reveal_col:
        reveal = st.button("👁️ اكشف الشرح", key=f"reveal_{pid}", use_container_width=True)
    with know_col:
        knew = st.button("عرفتها ✅", key=f"knew_{pid}", use_container_width=True)
    with unsure_col:
        unsure = st.button("متردد 🟡", key=f"unsure_{pid}", use_container_width=True)
    with wrong_col:
        wrong = st.button("ما عرفتها ❌", key=f"wrong_{pid}", use_container_width=True)

    if reveal or knew or unsure or wrong:
        st.session_state.atlas_quiz_revealed = True
        st.session_state.atlas_seen.add(pid)
        if knew:
            st.session_state.atlas_quiz_results[pid] = "known"
        elif unsure:
            st.session_state.atlas_quiz_results[pid] = "uncertain"
        elif wrong:
            st.session_state.atlas_quiz_results[pid] = "wrong"
            st.session_state.atlas_red_zone.add(pid)

    if st.session_state.get("atlas_quiz_revealed"):
        st.success(f"المحور المفهرس: {topic_label(page['topic_id'])}")
        st.caption(f"المصدر: {source['label']} · PDF page {page['pdf_page']} · {page['chapter']}")
        with st.expander("📄 اكشف نص الصفحة", expanded=True):
            st.text(page["text"] or "لا يوجد نص قابل للاستخراج من هذه الصفحة.")
        topic = TOPIC_BY_ID.get(page["topic_id"], TOPIC_BY_ID["other"])
        st.markdown("**Board pearls:**")
        for pearl in topic["pearls"]:
            st.markdown(f"- {pearl}")
        if pid in st.session_state.atlas_red_zone:
            st.warning("تمت إضافة الصورة إلى القبو الأحمر؛ ارجع لها في آخر جلسة.")


def render_mcq_tab(uploaded_files: list[Any]) -> None:
    st.header("📚 بنك الأسئلة — مكمل للأطلس")
    st.caption("هذا الوضع منفصل عن الأطلس. الإجابة غير الصريحة تظهر Unscored ولا تُحسب صحيحة.")
    if not uploaded_files:
        st.info("إذا تريد دمج بنك الأسئلة، ارفعه من الشريط الجانبي بصيغة PDF أو DOCX.")
        return

    questions: list[dict[str, Any]] = []
    warnings: list[str] = []
    for uploaded in uploaded_files:
        try:
            parsed = parse_mcq_file(uploaded.getvalue(), str(uploaded.name))
        except Exception as exc:
            parsed = []
            warnings.append(f"تعذر قراءة {uploaded.name}: {type(exc).__name__}")
        questions.extend(parsed)
        if not parsed:
            warnings.append(f"لم أجد أسئلة واضحة في {uploaded.name}.")
    for warning in warnings:
        st.warning(warning)
    if not questions:
        st.error("لم يتم استخراج أسئلة. إذا كان الملف صوراً، يحتاج OCR قبل رفعه.")
        return

    bank_key = hashlib.sha256(
        "|".join(f"{q['source']}|{q['q']}" for q in questions).encode("utf-8")
    ).hexdigest()[:16]
    if st.session_state.get("mcq_key") != bank_key:
        reset_mcq_state(bank_key)

    full_indices = list(range(len(questions)))
    if st.session_state.mcq_review_mode:
        current_indices = st.session_state.mcq_review_indices
    else:
        current_indices = full_indices
    if not current_indices:
        st.success("لا توجد أسئلة بهذا النوع من المراجعة.")
        return

    stats_col1, stats_col2, stats_col3 = st.columns(3)
    correct = sum(1 for i in full_indices if mcq_status(i) == "correct")
    wrong = sum(1 for i in full_indices if mcq_status(i) == "wrong")
    stats_col1.metric("صح", correct)
    stats_col2.metric("خطأ", wrong)
    stats_col3.metric("محلول", correct + wrong + sum(1 for i in full_indices if mcq_status(i) == "unscored"))

    if st.session_state.mcq_idx >= len(current_indices):
        st.success("وصلت إلى نهاية هذا القسم.")
        if st.button("↩️ العودة إلى كل الأسئلة", key="mcq_exit_review"):
            st.session_state.mcq_review_mode = False
            st.session_state.mcq_review_indices = []
            st.session_state.mcq_idx = 0
            st.session_state.mcq_revealed = False
            st.rerun()
        return

    position = st.session_state.mcq_idx
    original_index = current_indices[position]
    question = questions[original_index]
    current_status = mcq_status(original_index)
    st.markdown(f"### سؤال {position + 1} من {len(current_indices)}")
    st.caption(f"المصدر: {question['source']}")
    marked = st.checkbox("📌 تأشير السؤال", value=original_index in st.session_state.mcq_marks, key=f"mcq_mark_{bank_key}_{original_index}")
    if marked:
        st.session_state.mcq_marks.add(original_index)
    else:
        st.session_state.mcq_marks.discard(original_index)
    st.text(question["q"])
    choices = ["اختر إجابة...", "A", "B", "C", "D", "E"]
    choice = st.radio("اختيارك", choices, horizontal=True, key=f"mcq_choice_{bank_key}_{original_index}", disabled=st.session_state.mcq_revealed)
    if st.button("✅ تأكيد", key=f"mcq_confirm_{bank_key}_{original_index}", disabled=st.session_state.mcq_revealed):
        if choice == "اختر إجابة...":
            st.warning("اختَر جواباً أولاً.")
        else:
            status = "unscored" if question["letter"] == "?" else ("correct" if choice == question["letter"] else "wrong")
            st.session_state.mcq_results[str(original_index)] = status
            st.session_state.mcq_revealed = True
            if status == "wrong":
                st.session_state.atlas_red_zone.add(f"MCQ:{bank_key}:{original_index}")

    if st.session_state.mcq_revealed:
        status = mcq_status(original_index)
        if status == "correct":
            st.success(f"إجابة صحيحة ✅ — الجواب: {question['letter']}")
        elif status == "wrong":
            st.error(f"إجابة خاطئة ❌ — الجواب الصحيح: {question['letter']}")
        else:
            st.warning("لا يوجد حرف جواب صريح؛ لم تُحسب الإجابة صحيحة.")
        if question["expl"]:
            st.info(question["expl"])

    nav1, nav2, nav3 = st.columns(3)
    with nav1:
        if st.button("⬅️ السابق", disabled=position == 0, key=f"mcq_prev_{bank_key}_{original_index}"):
            st.session_state.mcq_idx -= 1
            st.session_state.mcq_revealed = False
            st.rerun()
    with nav2:
        if st.button("➡️ التالي", key=f"mcq_next_{bank_key}_{original_index}"):
            st.session_state.mcq_idx += 1
            st.session_state.mcq_revealed = False
            st.rerun()
    with nav3:
        mode = st.selectbox("مراجعة", ["كل الأسئلة", "الأخطاء", "المؤشرة"], key=f"mcq_review_mode_{bank_key}")
        if st.button("ابدأ المراجعة", key=f"mcq_start_review_{bank_key}"):
            mode_id = {"كل الأسئلة": "all", "الأخطاء": "wrong", "المؤشرة": "marked"}[mode]
            selected = mcq_review_indices(len(questions), mode_id)
            if selected:
                st.session_state.mcq_review_mode = True
                st.session_state.mcq_review_indices = selected
                st.session_state.mcq_idx = 0
                st.session_state.mcq_revealed = False
                st.rerun()
            st.warning("لا توجد أسئلة ضمن هذا الخيار.")

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Question", "Source", "Result", "Marked"])
    for i, item in enumerate(questions):
        writer.writerow([item["q"], item["source"], mcq_status(i) or "Not solved", "Yes" if i in st.session_state.mcq_marks else "No"])
    st.download_button("📥 تنزيل تقرير MCQ", output.getvalue().encode("utf-8-sig"), f"MCQ_progress_{today_iso()}.csv", "text/csv", key=f"mcq_download_{bank_key}")


def render_progress_tab(sources: list[dict[str, Any]]) -> None:
    st.header("📈 تقدمي")
    seen = len(st.session_state.get("atlas_seen", set()))
    marked = len(st.session_state.get("atlas_marks", set()))
    red = len(st.session_state.get("atlas_red_zone", set()))
    known = sum(1 for value in st.session_state.get("atlas_quiz_results", {}).values() if value == "known")
    uncertain = sum(1 for value in st.session_state.get("atlas_quiz_results", {}).values() if value == "uncertain")
    wrong = sum(1 for value in st.session_state.get("atlas_quiz_results", {}).values() if value == "wrong")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("صور مقروءة", seen)
    c2.metric("مؤشرة", marked)
    c3.metric("القبو الأحمر", red)
    c4.metric("اختبارات الصور", known + uncertain + wrong)
    st.caption(f"الاختبارات: {known} عرفتَها · {uncertain} متردد · {wrong} لم تعرفها")

    st.download_button(
        "📥 تنزيل تقدم الأطلس CSV",
        atlas_progress_csv(),
        f"Atlas_progress_{today_iso()}.csv",
        "text/csv",
        key="atlas_progress_csv",
    )
    st.download_button(
        "💾 تنزيل ملف التقدم الكامل JSON",
        json.dumps(progress_payload(), ensure_ascii=False, indent=2).encode("utf-8"),
        f"StudyHub_backup_{today_iso()}.json",
        "application/json",
        key="studyhub_backup_json",
    )
    restore = st.file_uploader("♻️ استرجاع JSON من جهاز آخر", type=["json"], key="studyhub_restore")
    if restore and st.button("استرجاع التقدم", key="restore_progress_button"):
        st.session_state.restore_notice = restore_progress(restore.getvalue())
        st.rerun()
    if st.session_state.get("restore_notice"):
        st.success(st.session_state.restore_notice)
        st.session_state.restore_notice = ""

    st.markdown("### لماذا لا يوجد JSONBin افتراضياً؟")
    st.write("حتى لا يتعطل التطبيق إذا نسيت Secrets، وحتى لا تُرسل بياناتك إلى خدمة خارجية. استخدم JSON backup بين الهاتف واللابتوب إذا احتجت.")
    if st.button("🧹 تصفير تقدم هذه الجلسة", key="reset_all_progress"):
        st.session_state.atlas_marks = set()
        st.session_state.atlas_red_zone = set()
        st.session_state.atlas_seen = set()
        st.session_state.atlas_quiz_results = {}
        st.session_state.mcq_results = {}
        st.session_state.mcq_marks = set()
        st.success("تم تصفير التقدم داخل الجلسة الحالية فقط.")


st.title("🚀 Pediatric Hem/Onc Atlas Study Hub 🩺")
st.caption(f"نسخة {APP_VERSION} · كتب الأطلس تُرفع داخل الجلسة ولا تُضمّن في المستودع")
st.info("🔒 خلي المستودع والتطبيق Private. لا ترفع ملفات الكتب إلى GitHub؛ ارفعها من داخل التطبيق عند الدراسة.")

st.sidebar.markdown("## 📖 ملفات الأطلس")
atlas_uploads = st.sidebar.file_uploader(
    "ارفع Rodak 5th وHoffbrand 4th (اختياري Rodak 4th)",
    type=["pdf"],
    accept_multiple_files=True,
    key="atlas_uploads",
)
st.sidebar.caption("الأفضل: Rodak الخامس + Hoffbrand. النسخة العربية المكررة أقدم، فلا تحتاجها غالباً.")

st.sidebar.markdown("## 📚 ملفات MCQ (اختياري)")
mcq_uploads = st.sidebar.file_uploader(
    "ارفع PDF/DOCX لأسئلة البورد",
    type=["pdf", "docx"],
    accept_multiple_files=True,
    key="mcq_uploads",
)

sources: list[dict[str, Any]] = []
if atlas_uploads:
    with st.spinner("جاري فهرسة عناوين وصفحات الأطلس…"):
        sources = [make_source(uploaded) for uploaded in atlas_uploads]

tab_plan, tab_atlas, tab_quiz, tab_mcq, tab_progress = st.tabs(
    ["🧭 خطة البورد", "🔬 مستكشف الأطلس", "🧠 اختبار الصور", "📚 MCQ", "📈 تقدمي"]
)

with tab_plan:
    render_plan(sources)
with tab_atlas:
    render_atlas_explorer(sources)
with tab_quiz:
    render_image_quiz(sources)
with tab_mcq:
    render_mcq_tab(mcq_uploads or [])
with tab_progress:
    render_progress_tab(sources)
