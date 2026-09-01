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
except ImportError as exc:  # pragma: no cover
    raise RuntimeError("PyMuPDF is required. Install requirements.txt first.") from exc


APP_TITLE = "Pediatric Hem/Onc Atlas Exam Hub"
APP_VERSION = "2.0"

st.set_page_config(page_title=APP_TITLE, layout="wide")


# -----------------------------------------------------------------------------
# Atlas slots
# -----------------------------------------------------------------------------
# Each slot has its own upload field and its own exam crop profile. This is
# intentionally explicit: filenames such as 1.pdf and 2.pdf are not reliable
# enough to decide which atlas they contain.
ATLAS_SPECS: list[dict[str, Any]] = [
    {
        "id": "rodak6",
        "title": "Rodak Clinical Hematology Atlas — Sixth Edition (2022)",
        "hint": "الملف الذي ظهر عندنا باسم 3.pdf",
        "profile": "rodak_fullpage",
        "recommended": True,
        "crop_note": "صفحات scan كاملة؛ القص يعتمد على صفوف الصور قبل الـ figure caption.",
    },
    {
        "id": "rodak5",
        "title": "Rodak Clinical Hematology Atlas — Fifth Edition (2017)",
        "hint": "book Dr Maha(1).pdf",
        "profile": "figure_blocks",
        "recommended": True,
        "crop_note": "يأخذ مربعات الصور نفسها بدون المساحة العمودية التي تحتوي الكابشن.",
    },
    {
        "id": "wintrobe_atlas",
        "title": "Wintrobe's Atlas of Clinical Hematology — Second Edition (2018)",
        "hint": "الملف الذي ظهر عندنا باسم 2.pdf",
        "profile": "figure_blocks",
        "recommended": True,
        "crop_note": "يستخرج figure blocks الكبيرة ويترك نص الشرح والكابشن خارج القص.",
    },
    {
        "id": "anderson",
        "title": "Anderson's Atlas of Hematology — Second Edition (2014)",
        "hint": "الملف الذي ظهر عندنا باسم 4.pdf",
        "profile": "anderson",
        "recommended": True,
        "crop_note": "يستبعد شريط العنوان الصغير ويأخذ الصورة الرئيسية فقط.",
    },
    {
        "id": "hoffbrand",
        "title": "Color Atlas of Clinical Hematology — Fourth Edition (2010)",
        "hint": "color atlas of clinical hematology...pdf",
        "profile": "hoffbrand",
        "recommended": True,
        "crop_note": "يعرض panels الصور منفصلة حتى لا يظهر نص Fig. أو الوصف.",
    },
    {
        "id": "rodak4",
        "title": "Rodak Clinical Hematology Atlas — Fourth Edition (2013)",
        "hint": "6.pdf أو رهيب ويه كتاب DR MAHA(1).pdf",
        "profile": "figure_blocks",
        "recommended": False,
        "crop_note": "نفس أسلوب Rodak: الصورة exact بدون vertical padding.",
    },
]

REFERENCE_HINT = "1.pdf = Wintrobe's Clinical Hematology 14th Edition؛ كتاب مرجعي، وليس أطلس صور."

TOPICS: list[dict[str, Any]] = [
    {
        "id": "foundation",
        "label": "Foundation: blood film and maturation",
        "keywords": [
            "hematopoiesis", "maturation", "peripheral blood", "bone marrow", "blood film",
            "blood smear", "cell machinery", "growth factors", "microscopic evaluation",
        ],
        "pearls": [
            "ابدأ بالوصف: RBC ثم WBC ثم platelets ثم الخلفية.",
            "حدد مرحلة النضج من شكل النواة والكروماتين والنويات والحبيبات.",
            "لا تجعل شكل خلية واحدة يلغي pattern الصورة كلها.",
        ],
    },
    {
        "id": "rbc_morphology",
        "label": "RBC morphology and inclusions",
        "keywords": [
            "erythrocyte", "red cell", "anisocytosis", "poikilocytosis", "microcyte",
            "macrocyte", "hypochromia", "target cell", "spherocyte", "schistocyte",
            "elliptocyte", "ovalocyte", "teardrop", "tear drop", "rouleaux", "inclusion",
            "basophilic stippling", "howell-jolly", "pappenheimer", "polychrom",
        ],
        "pearls": [
            "Schistocytes تعني RBC fragmentation؛ اربطها بسياق MAHA/DIC/TMA.",
            "Spherocytes بلا central pallor؛ فرّق HS عن immune hemolysis بالسياق وDAT.",
            "Target cells ليست تشخيصاً وحدها؛ اربطها بالهيموغلوبينوباثي أو الكبد أو نقص الطحال.",
        ],
    },
    {
        "id": "anemia",
        "label": "Anemia and hemoglobin disorders",
        "keywords": [
            "anemia", "anaemia", "hypochromic", "hemolytic", "megaloblastic", "aplastic",
            "thalassemia", "sickle", "hemoglobin", "hemoglobinopathy", "iron deficiency",
            "g6pd", "glucose-6-phosphate", "hereditary spherocytosis", "porphyria",
            "dyserythropoietic", "iron overload",
        ],
        "pearls": [
            "Microcytosis لا تساوي iron deficiency؛ قارن RBC count وRDW والـ iron studies.",
            "في hemolysis اجمع morphology مع reticulocytes وbilirubin وLDH وDAT.",
            "في sickle/thalassemia اربط smear مع electrophoresis أو genotype، لا تعتمد على الصورة وحدها.",
        ],
    },
    {
        "id": "acute_leukemia",
        "label": "Acute leukemia: ALL / AML / APL",
        "keywords": [
            "acute leukemia", "acute leukemias", "acute myeloid", "acute lymphoblastic", "all",
            "aml", "blast", "lymphoblast", "myeloblast", "promyelocytic", "auer rod",
            "peroxidase", "myeloperoxidase", "monoblast", "megakaryoblast",
        ],
        "pearls": [
            "Auer rods تدعم myeloid differentiation؛ bundles داخل abnormal promyelocytes ترفع الشك بـ APL.",
            "MPO وflow cytometry وgenetics تكمل morphology؛ لا تحدد lineage من الصورة وحدها.",
            "الأطالس القديمة قد تستعمل تسميات قديمة؛ راجع WHO/ICC عند سؤال التصنيف.",
        ],
    },
    {
        "id": "myeloid_lymphoid",
        "label": "MDS / MPN / lymphoproliferative disorders",
        "keywords": [
            "myelodysplastic", "mds", "myeloproliferative", "mpn", "chronic myeloid", "cml",
            "chronic lymphoid", "lymphoproliferative", "non-hodgkin", "hodgkin", "myeloma",
            "dysplasia", "polycythemia", "thrombocythemia", "myelofibrosis",
        ],
        "pearls": [
            "Left shift وحده لا يساوي CML؛ اربطه بالـ basophilia والـ molecular context.",
            "Dysplasia هي pattern متكرر، وليست خلية شاذة منفردة.",
            "تأكد من الاسم والتصنيف الحديث قبل اعتماد مصطلح قديم من أطلس أقدم.",
        ],
    },
    {
        "id": "platelet_coagulation",
        "label": "Platelets, hemostasis and coagulation",
        "keywords": [
            "platelet", "thrombocytopenia", "thrombocytosis", "hemostasis", "bleeding",
            "coagulation", "vascular", "purpura", "thrombosis", "megakaryocyte", "von willebrand",
        ],
        "pearls": [
            "حجم الصفائح ووجود giant platelets يوجهانك، لكن لا يكفيان وحدهما للتشخيص.",
            "اربط smear مع platelet count وPT/aPTT وfibrinogen والقصة السريرية.",
            "لا تستخدم morphology وحدها لنفي DIC أو TMA.",
        ],
    },
    {
        "id": "histiocytic_infection",
        "label": "Histiocytic disorders and parasites",
        "keywords": [
            "histiocytic", "histiocyte", "langerhans", "lch", "hemophagocytosis", "parasite",
            "malaria", "babesia", "leishmania", "microorganism", "infection", "storage disorder",
        ],
        "pearls": [
            "Hemophagocytosis داعمة لكنها ليست وحدها كافية لتشخيص HLH.",
            "LCH تحتاج ربط morphology بالـ immunophenotype والسياق السريري.",
            "الطفيلي يحتاج الشكل والمكان والسياق الوبائي، لا تشابه بصري فقط.",
        ],
    },
    {
        "id": "other",
        "label": "Other: newborn, fluids, transfusion and miscellaneous",
        "keywords": [
            "newborn", "body fluids", "transfusion", "stem cell transplantation", "transplantation",
            "miscellaneous", "nonhematopoietic", "plasma cell", "cell descriptions",
        ],
        "pearls": [
            "ضع هذه المجموعة بعد تثبيت RBC واللوكيميا والـ marrow disorders.",
            "أي صورة تتكرر في بنك أسئلتك ارفعها إلى القبو الأحمر حتى لو كانت من محور ثانوي.",
        ],
    },
]
TOPIC_BY_ID = {topic["id"]: topic for topic in TOPICS}


def init_state() -> None:
    defaults: dict[str, Any] = {
        "atlas_marks": set(),
        "atlas_red_zone": set(),
        "atlas_seen": set(),
        "atlas_page_meta": {},
        "atlas_quiz_results": {},
        "atlas_quiz_page": None,
        "atlas_quiz_signature": None,
        "atlas_quiz_revealed": False,
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


def clean_text(text: str) -> str:
    # Some scans contain isolated surrogate characters in bookmark strings.
    return re.sub(r"\s+", " ", str(text or "").encode("utf-8", "ignore").decode("utf-8", "ignore")).strip()


def safe_int(value: Any, default: int = 0) -> int:
    try:
        return int(str(value).strip())
    except (TypeError, ValueError):
        return default


def topic_label(topic_id: str) -> str:
    return TOPIC_BY_ID.get(topic_id, TOPIC_BY_ID["other"])["label"]


def infer_topic(chapter: str, text: str) -> str:
    chapter_lower = chapter.lower()
    chapter_rules = [
        ("acute_leukemia", ["acute leukemia", "acute leukemias", "acute myeloid", "precursor lymphoid"]),
        ("histiocytic_infection", ["histiocytic", "microorganism", "parasitic", "parasite"]),
        ("platelet_coagulation", ["platelet", "hemostasis", "bleeding disorders", "coagulation", "thrombosis"]),
        ("myeloid_lymphoid", [
            "myelodysplastic", "myeloproliferative", "chronic myeloid", "chronic lymphoid",
            "lymphoproliferative", "non-hodgkin", "hodgkin lymphoma", "myeloma",
        ]),
        ("anemia", [
            "hypochromic anemia", "hemolytic anemia", "megaloblastic", "aplastic",
            "dyserythropoietic", "porphyria", "iron overload", "genetic disorders of hemoglobin",
            "diseases affecting erythrocytes", "anemia",
        ]),
        ("rbc_morphology", [
            "variations in size and color of erythrocytes", "variations in shape and distribution of erythrocytes",
            "inclusions in erythrocytes", "erythrocyte maturation", "red blood cells", "erythrocytes",
        ]),
        ("foundation", [
            "hematopoiesis", "cellular basis of hematopoiesis", "cell machinery", "growth factors",
            "maturation of blood cells", "peripheral blood film", "blood film examination",
            "microscopic evaluation", "blood cells",
        ]),
    ]
    for topic_id, phrases in chapter_rules:
        if any(phrase in chapter_lower for phrase in phrases):
            return topic_id

    haystack = f"{chapter} {text}".lower()
    best_id, best_score = "other", 0
    for topic in TOPICS:
        if topic["id"] == "other":
            continue
        score = sum(1 + haystack.count(keyword.lower()) // 12 for keyword in topic["keywords"] if keyword.lower() in haystack)
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


@st.cache_data(show_spinner=False, max_entries=20)
def build_pdf_index(file_bytes: bytes, atlas_id: str) -> dict[str, Any]:
    """Build text/bookmark metadata; images are rendered only for the page in use."""
    document = fitz.open(stream=file_bytes, filetype="pdf")
    try:
        raw_toc = document.get_toc(simple=True) or []
    except Exception:
        raw_toc = []

    chapters: list[dict[str, Any]] = []
    for item in raw_toc:
        if len(item) < 3:
            continue
        page_number = safe_int(item[2], 0)
        title = clean_text(item[1])
        if page_number >= 1 and title:
            chapters.append({"pdf_page": page_number, "title": title, "level": safe_int(item[0], 1)})
    chapters.sort(key=lambda row: (row["pdf_page"], row["level"]))

    pages: list[dict[str, Any]] = []
    for page_index in range(len(document)):
        page = document.load_page(page_index)
        page_number = page_index + 1
        text = clean_text(page.get_text("text"))
        chapter = chapter_for_page(chapters, page_number)
        topic_id = infer_topic(chapter, text)
        is_cover = bool(re.search(r"\bchapter\s+\d+\b", text.lower())) and len(text) < 240
        pages.append(
            {
                "pdf_page": page_number,
                "chapter": chapter,
                "topic_id": topic_id,
                "text": text[:9000],
                "has_text": bool(text),
                "has_images": bool(page.get_images(full=True)),
                "is_cover": is_cover,
            }
        )
    document.close()
    return {"atlas_id": atlas_id, "page_count": len(pages), "chapters": chapters, "pages": pages}


def make_source(spec: dict[str, Any], uploaded: Any) -> dict[str, Any]:
    data = uploaded.getvalue()
    digest = hashlib.sha256(data).hexdigest()[:16]
    return {
        "id": f"{spec['id']}:{digest}",
        "slot_id": spec["id"],
        "title": spec["title"],
        "filename": str(uploaded.name),
        "bytes": data,
        "profile": spec["profile"],
        "recommended": spec["recommended"],
        "crop_note": spec["crop_note"],
        "index": build_pdf_index(data, spec["id"]),
    }


def source_by_id(sources: list[dict[str, Any]], source_id: str) -> dict[str, Any] | None:
    return next((source for source in sources if source["id"] == source_id), None)


def page_by_number(source: dict[str, Any], pdf_page: int) -> dict[str, Any] | None:
    return next((page for page in source["index"]["pages"] if page["pdf_page"] == pdf_page), None)


def page_key(source: dict[str, Any], page: dict[str, Any]) -> str:
    return f"{source['id']}:{page['pdf_page']}"


def remember_page(source: dict[str, Any], page: dict[str, Any]) -> str:
    pid = page_key(source, page)
    st.session_state.atlas_page_meta[pid] = {
        "source": source["title"],
        "file": source["filename"],
        "pdf_page": page["pdf_page"],
        "topic": topic_label(page["topic_id"]),
        "chapter": page["chapter"],
    }
    return pid


def usable_sources(sources: list[dict[str, Any]], include_old: bool) -> list[dict[str, Any]]:
    if include_old:
        return sources
    preferred = [source for source in sources if source["slot_id"] != "rodak4"]
    return preferred or sources


def page_matches(page: dict[str, Any], topic_id: str, search: str) -> bool:
    if topic_id != "all" and page["topic_id"] != topic_id:
        return False
    if search:
        haystack = f"{page['chapter']} {page['text']}".lower()
        return search.lower().strip() in haystack
    return True


def _image_rects(page: Any, profile: str) -> list[Any]:
    blocks = page.get_text("dict").get("blocks", [])
    page_area = page.rect.width * page.rect.height
    min_fraction = {
        "anderson": 0.018,
        "hoffbrand": 0.012,
        "figure_blocks": 0.010,
        "rodak_fullpage": 0.010,
    }.get(profile, 0.012)
    candidates: list[tuple[float, Any]] = []
    for block in blocks:
        if block.get("type") != 1 or not block.get("bbox"):
            continue
        x0, y0, x1, y1 = block["bbox"]
        rect = fitz.Rect(x0, y0, x1, y1)
        area = max(0.0, rect.width) * max(0.0, rect.height)
        if area >= page_area * min_fraction and rect.width > 30 and rect.height > 30:
            candidates.append((area, rect))

    # Remove nested/duplicate rectangles and arrange panels as they appear.
    candidates.sort(key=lambda pair: pair[0], reverse=True)
    selected: list[Any] = []
    for _, rect in candidates:
        duplicate = False
        for other in selected:
            overlap = rect & other
            if overlap.get_area() >= rect.get_area() * 0.88:
                duplicate = True
                break
        if not duplicate:
            selected.append(rect)
    selected.sort(key=lambda rect: (round(rect.y0 / 8), rect.x0))

    if profile == "anderson" and len(selected) > 1:
        # Anderson pages often have a tiny decorative header image followed by
        # one large hematology photograph.
        selected = [max(selected, key=lambda rect: rect.get_area())]
    return selected[:12]


def _caption_rects(page: Any) -> list[Any]:
    captions: list[Any] = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0 or not block.get("bbox"):
            continue
        text = clean_text(" ".join(
            span.get("text", "")
            for line in block.get("lines", [])
            for span in line.get("spans", [])
        ))
        if re.match(r"(?i)^(?:fig(?:ure)?\.?|figure)\s*[A-Z0-9]", text):
            captions.append(fitz.Rect(*block["bbox"]))
    return sorted(captions, key=lambda rect: rect.y0)


def _panel_label_rects(page: Any) -> list[Any]:
    """Find standalone A/B/C panel labels, not letters inside captions."""
    labels: list[Any] = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0 or not block.get("bbox"):
            continue
        text = clean_text(" ".join(
            span.get("text", "")
            for line in block.get("lines", [])
            for span in line.get("spans", [])
        ))
        if re.fullmatch(r"[A-Z](?:\s+[A-Z])*", text) and len(text) <= 15:
            labels.append(fitz.Rect(*block["bbox"]))
    return labels


def _trim_panel_labels(page: Any, rects: list[Any]) -> list[Any]:
    """Trim the small panel-letter strip when it is outside the photograph."""
    labels = _panel_label_rects(page)
    trimmed: list[Any] = []
    for rect in rects:
        relevant = [
            label for label in labels
            if label.y0 >= rect.y0 - 2
            and label.y0 < rect.y1
            and label.x1 > rect.x0
            and label.x0 < rect.x1
        ]
        if relevant:
            new_bottom = min(label.y0 for label in relevant) - 3
            if new_bottom > rect.y0 + rect.height * 0.55:
                rect = fitz.Rect(rect.x0, rect.y0, rect.x1, new_bottom)
        trimmed.append(rect)
    return trimmed


def _fullpage_figure_rects(page: Any, profile: str) -> list[Any]:
    """For full-page scans, use the space above each caption, never the caption itself."""
    captions = _caption_rects(page)
    if captions:
        first_caption_y = captions[0].y0
        top_text_bottom = page.rect.y0 + page.rect.height * 0.10
        for block in page.get_text("dict").get("blocks", []):
            if block.get("type") != 0 or not block.get("bbox"):
                continue
            rect = fitz.Rect(*block["bbox"])
            if rect.y1 < first_caption_y and rect.y0 < page.rect.y0 + page.rect.height * 0.25:
                top_text_bottom = max(top_text_bottom, rect.y1)

        rects: list[Any] = []
        for index, caption in enumerate(captions):
            start = top_text_bottom + 6 if index == 0 else captions[index - 1].y1 + 10
            end = caption.y0 - 4
            if end - start > page.rect.height * 0.07:
                rects.append(fitz.Rect(page.rect.x0, start, page.rect.x1, end))
        if rects:
            return _trim_panel_labels(page, rects[:8])

    # Last resort for pages without a text-layer caption. This is intentionally
    # conservative: it removes the header and footer but cannot remove labels
    # baked into a raster image.
    if profile == "rodak_fullpage":
        rects = [fitz.Rect(
            page.rect.x0,
            page.rect.y0 + page.rect.height * 0.16,
            page.rect.x1,
            page.rect.y1 - page.rect.height * 0.16,
        )]
        return _trim_panel_labels(page, rects)
    rects = [fitz.Rect(
        page.rect.x0,
        page.rect.y0 + page.rect.height * 0.10,
        page.rect.x1,
        page.rect.y1 - page.rect.height * 0.12,
    )]
    return _trim_panel_labels(page, rects)


@st.cache_data(show_spinner=False, max_entries=100)
def render_exam_parts(file_bytes: bytes, pdf_page: int, profile: str, zoom: float) -> list[bytes]:
    document = fitz.open(stream=file_bytes, filetype="pdf")
    page = document.load_page(max(0, pdf_page - 1))
    rects = _image_rects(page, profile)
    page_area = page.rect.width * page.rect.height
    has_fullpage_image = any(rect.get_area() >= page_area * 0.86 for rect in rects)
    if profile == "rodak_fullpage" or has_fullpage_image or not rects:
        rects = _fullpage_figure_rects(page, profile)
    else:
        rects = _trim_panel_labels(page, rects)

    parts: list[bytes] = []
    for rect in rects:
        pixmap = page.get_pixmap(matrix=fitz.Matrix(float(zoom), float(zoom)), clip=rect, alpha=False)
        parts.append(pixmap.tobytes("png"))
    document.close()
    return parts


@st.cache_data(show_spinner=False, max_entries=80)
def render_full_page(file_bytes: bytes, pdf_page: int, zoom: float) -> bytes:
    document = fitz.open(stream=file_bytes, filetype="pdf")
    page = document.load_page(max(0, pdf_page - 1))
    pixmap = page.get_pixmap(matrix=fitz.Matrix(float(zoom), float(zoom)), alpha=False)
    result = pixmap.tobytes("png")
    document.close()
    return result


def choose_exam_page(sources: list[dict[str, Any]], topic_id: str, source_id: str) -> dict[str, Any] | None:
    candidates: list[dict[str, Any]] = []
    selected_sources = sources if source_id == "all" else [source for source in sources if source["id"] == source_id]
    for source in selected_sources:
        for page in source["index"]["pages"]:
            if topic_id != "all" and page["topic_id"] != topic_id:
                continue
            if page["is_cover"] or (not page["has_text"] and not page["has_images"]):
                continue
            candidates.append({"source_id": source["id"], "pdf_page": page["pdf_page"]})
    if not candidates:
        return None
    previous = st.session_state.get("atlas_quiz_page")
    alternatives = [item for item in candidates if item != previous] or candidates
    return random.choice(alternatives)


def progress_csv() -> bytes:
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Page ID", "Atlas", "PDF page", "Topic", "Status", "Marked", "Red zone"])
    keys = set(st.session_state.atlas_page_meta)
    keys |= set(st.session_state.atlas_seen)
    keys |= set(st.session_state.atlas_marks)
    keys |= set(st.session_state.atlas_red_zone)
    for pid in sorted(keys):
        meta = st.session_state.atlas_page_meta.get(pid, {})
        writer.writerow([
            pid,
            meta.get("source", ""),
            meta.get("pdf_page", ""),
            meta.get("topic", ""),
            st.session_state.atlas_quiz_results.get(pid, "seen"),
            "Yes" if pid in st.session_state.atlas_marks else "No",
            "Yes" if pid in st.session_state.atlas_red_zone else "No",
        ])
    return output.getvalue().encode("utf-8-sig")


def backup_payload() -> dict[str, Any]:
    return {
        "app": APP_TITLE,
        "version": APP_VERSION,
        "saved_at": dt.datetime.now().isoformat(timespec="seconds"),
        "atlas_marks": sorted(st.session_state.atlas_marks),
        "atlas_red_zone": sorted(st.session_state.atlas_red_zone),
        "atlas_seen": sorted(st.session_state.atlas_seen),
        "atlas_page_meta": dict(st.session_state.atlas_page_meta),
        "atlas_quiz_results": dict(st.session_state.atlas_quiz_results),
    }


def restore_backup(raw: bytes) -> str:
    try:
        data = json.loads(raw.decode("utf-8"))
        if not isinstance(data, dict):
            return "ملف التقدم غير صالح."
        for key in ("atlas_marks", "atlas_red_zone", "atlas_seen"):
            if isinstance(data.get(key), list):
                st.session_state[key] = set(str(value) for value in data[key])
        if isinstance(data.get("atlas_page_meta"), dict):
            st.session_state.atlas_page_meta = data["atlas_page_meta"]
        if isinstance(data.get("atlas_quiz_results"), dict):
            st.session_state.atlas_quiz_results = {str(k): str(v) for k, v in data["atlas_quiz_results"].items()}
        return "تم استرجاع تقدم الأطلس."
    except (UnicodeDecodeError, json.JSONDecodeError, TypeError):
        return "تعذر قراءة ملف التقدم."


# -----------------------------------------------------------------------------
# UI sections
# -----------------------------------------------------------------------------
def render_upload_tab() -> list[dict[str, Any]]:
    st.header("📥 تحميل كل أطلس في حقله")
    st.write("ارفع كل ملف في الحقل المطابق له. لا ترفع ملفات PDF إلى GitHub؛ ارفعها هنا داخل التطبيق.")
    st.info("النسخة الجديدة أطالس فقط؛ لا يوجد MCQ ولا تصحيح خيارات.")

    uploads: list[dict[str, Any]] = []
    for spec in ATLAS_SPECS:
        expanded = spec["recommended"] and spec["id"] in {"rodak5", "rodak6"}
        with st.expander(f"📕 {spec['title']}", expanded=expanded):
            st.caption(spec["hint"])
            uploaded = st.file_uploader(
                "ارفع ملف PDF لهذا الأطلس",
                type=["pdf"],
                accept_multiple_files=False,
                key=f"upload_{spec['id']}",
            )
            if uploaded:
                uploads.append({"spec": spec, "uploaded": uploaded})
                st.success(f"تم اختيار: {uploaded.name}")

    with st.expander("📘 Wintrobe Clinical Hematology 14th — مرجع نصي فقط", expanded=False):
        st.caption(REFERENCE_HINT)
        st.warning("لا يدخل وضع امتحان الصور، وحجمه كبير جداً؛ لا ترفعه إلا إذا تحتاج البحث النصي لاحقاً.")
        st.file_uploader("رفع اختياري — لا يدخل الفهرسة", type=["pdf"], key="reference_upload")

    if uploads:
        with st.spinner("جاري فهرسة عناوين الأطالس…"):
            sources = [make_source(item["spec"], item["uploaded"]) for item in uploads]
        st.success(f"جاهز {len(sources)} أطلس/أطالس للاستخدام.")
        for source in sources:
            st.caption(f"{source['title']} · {source['index']['page_count']} صفحة · {source['crop_note']}")
        return sources

    st.warning("ارفع Rodak 5th أو Rodak 6th أولاً حتى يظهر محتوى الأطلس.")
    return []


def render_explorer_tab(sources: list[dict[str, Any]]) -> None:
    st.header("🔬 مستكشف الأطالس")
    if not sources:
        st.info("اذهب إلى تبويب تحميل الأطالس وارفع الملفات.")
        return

    source_ids = [source["id"] for source in sources]
    source_id = st.selectbox(
        "اختر الأطلس",
        source_ids,
        format_func=lambda sid: next(source["title"] for source in sources if source["id"] == sid),
        key="explorer_source",
    )
    source = source_by_id(sources, source_id)
    if not source:
        return

    topic_id = st.selectbox(
        "المحور",
        ["all"] + [topic["id"] for topic in TOPICS],
        format_func=lambda tid: "كل الصفحات" if tid == "all" else topic_label(tid),
        key="explorer_topic",
    )
    search = st.text_input("🔎 ابحث بكلمة English من عنوان المرض أو الخلية", key="explorer_search")
    matching = [page for page in source["index"]["pages"] if page_matches(page, topic_id, search)]
    st.caption(f"وجدت {len(matching)} صفحة من أصل {source['index']['page_count']}.")
    if not matching:
        st.warning("لا توجد صفحات بهذا الفلتر. جرّب كل الصفحات أو كلمة أبسط.")
        return

    page_numbers = [page["pdf_page"] for page in matching]
    selected_page_number = st.selectbox(
        "اختر صفحة PDF",
        page_numbers,
        format_func=lambda number: next(
            f"PDF page {number} · {page['chapter'][:80]}"
            for page in matching if page["pdf_page"] == number
        ),
        key=f"explorer_page_{source_id}_{topic_id}_{hash(search)}",
    )
    page = page_by_number(source, selected_page_number)
    if not page:
        return
    pid = remember_page(source, page)
    st.markdown(f"### {topic_label(page['topic_id'])}")
    st.caption(f"{source['title']} · PDF page {selected_page_number} · {page['chapter']}")
    zoom = st.slider("التكبير", 1.0, 2.2, 1.35, 0.05, key=f"explorer_zoom_{pid}")
    st.image(render_full_page(source["bytes"], selected_page_number, zoom), width="stretch")

    c1, c2, c3 = st.columns(3)
    with c1:
        marked = st.checkbox("📌 تأشير الصفحة", value=pid in st.session_state.atlas_marks, key=f"explorer_mark_{pid}")
        if marked:
            st.session_state.atlas_marks.add(pid)
        else:
            st.session_state.atlas_marks.discard(pid)
    with c2:
        red = st.checkbox("🔥 القبو الأحمر", value=pid in st.session_state.atlas_red_zone, key=f"explorer_red_{pid}")
        if red:
            st.session_state.atlas_red_zone.add(pid)
        else:
            st.session_state.atlas_red_zone.discard(pid)
    with c3:
        if st.button("✅ سجل كمقروءة", key=f"explorer_seen_{pid}", use_container_width=True):
            st.session_state.atlas_seen.add(pid)
            st.toast("تم تسجيل الصفحة.")

    with st.expander("📄 النص المستخرج من الصفحة"):
        st.text(page["text"] or "لا يوجد نص قابل للاستخراج؛ اعتمد على الصورة.")
    with st.expander("🩺 Board pearls لهذا المحور", expanded=True):
        for pearl in TOPIC_BY_ID.get(page["topic_id"], TOPIC_BY_ID["other"])["pearls"]:
            st.markdown(f"- {pearl}")


def render_exam_tab(sources: list[dict[str, Any]]) -> None:
    st.header("🧠 وضع الامتحان — صورة بلا عنوان أو كابشن")
    if not sources:
        st.info("ارفع الأطالس أولاً من تبويب التحميل.")
        return

    include_old = st.checkbox("إدخال Rodak 4th القديم ضمن الاختبار", value=False, key="exam_include_old")
    usable = usable_sources(sources, include_old)
    source_options = ["all"] + [source["id"] for source in usable]
    selected_source = st.selectbox(
        "الأطلس في الاختبار",
        source_options,
        format_func=lambda sid: "كل الأطالس المختارة" if sid == "all" else next(source["title"] for source in usable if source["id"] == sid),
        key="exam_source",
    )
    selected_topic = st.selectbox(
        "محور الصورة",
        ["all"] + [topic["id"] for topic in TOPICS if topic["id"] != "other"],
        format_func=lambda tid: "عشوائي من كل المحاور" if tid == "all" else topic_label(tid),
        key="exam_topic",
    )

    signature = f"{selected_source}:{selected_topic}:{include_old}"
    if st.session_state.atlas_quiz_signature != signature:
        st.session_state.atlas_quiz_signature = signature
        st.session_state.atlas_quiz_page = choose_exam_page(usable, selected_topic, selected_source)
        st.session_state.atlas_quiz_revealed = False

    if st.button("🎲 صورة امتحان جديدة", type="primary", key="exam_new_image"):
        st.session_state.atlas_quiz_page = choose_exam_page(usable, selected_topic, selected_source)
        st.session_state.atlas_quiz_revealed = False
        st.rerun()

    selected = st.session_state.atlas_quiz_page
    if not selected:
        st.warning("لم أجد صوراً مناسبة لهذا الفلتر.")
        return
    source = source_by_id(usable, selected["source_id"])
    page = page_by_number(source, selected["pdf_page"]) if source else None
    if not source or not page:
        st.warning("تعذر فتح الصورة؛ اضغط صورة امتحان جديدة.")
        return
    pid = remember_page(source, page)

    st.caption("لا يظهر اسم الأطلس أو رقم الصفحة قبل الكشف. اكتب وصفك أولاً.")
    parts = render_exam_parts(source["bytes"], page["pdf_page"], source["profile"], 1.45)
    if len(parts) == 1:
        st.image(parts[0], width="stretch")
    else:
        cols = st.columns(min(3, len(parts)))
        for index, part in enumerate(parts):
            with cols[index % len(cols)]:
                st.image(part, width="stretch")

    st.markdown("### قبل الكشف")
    st.write("اكتب 2–3 ملاحظات مورفولوجية، ثم differential أو التشخيص الأقوى. لا تقرأ العنوان أولاً.")
    st.text_area("وصفك", key=f"exam_guess_{pid}", height=115)

    c1, c2, c3, c4 = st.columns(4)
    reveal = c1.button("👁️ اكشف", key=f"exam_reveal_{pid}", use_container_width=True)
    known = c2.button("عرفتها ✅", key=f"exam_known_{pid}", use_container_width=True)
    uncertain = c3.button("متردد 🟡", key=f"exam_uncertain_{pid}", use_container_width=True)
    wrong = c4.button("ما عرفتها ❌", key=f"exam_wrong_{pid}", use_container_width=True)
    if reveal or known or uncertain or wrong:
        st.session_state.atlas_quiz_revealed = True
        st.session_state.atlas_seen.add(pid)
        if known:
            st.session_state.atlas_quiz_results[pid] = "known"
        elif uncertain:
            st.session_state.atlas_quiz_results[pid] = "uncertain"
        elif wrong:
            st.session_state.atlas_quiz_results[pid] = "wrong"
            st.session_state.atlas_red_zone.add(pid)

    if st.session_state.atlas_quiz_revealed:
        st.success(f"المحور المفهرس: {topic_label(page['topic_id'])}")
        st.caption(f"المصدر: {source['title']} · PDF page {page['pdf_page']} · {page['chapter']}")
        st.caption(f"طريقة القص الخاصة بهذا الأطلس: {source['crop_note']}")
        with st.expander("📄 اكشف نص الصفحة والكابشن", expanded=True):
            st.text(page["text"] or "لا يوجد نص قابل للاستخراج.")
        st.markdown("**Board pearls:**")
        for pearl in TOPIC_BY_ID.get(page["topic_id"], TOPIC_BY_ID["other"])["pearls"]:
            st.markdown(f"- {pearl}")
        if pid in st.session_state.atlas_red_zone:
            st.warning("أضيفت إلى القبو الأحمر.")
        if st.button("➡️ انتقل للصورة التالية", key=f"exam_next_{pid}"):
            st.session_state.atlas_quiz_page = choose_exam_page(usable, selected_topic, selected_source)
            st.session_state.atlas_quiz_revealed = False
            st.rerun()


def render_progress_tab() -> None:
    st.header("📈 تقدم الأطلس")
    known = sum(value == "known" for value in st.session_state.atlas_quiz_results.values())
    uncertain = sum(value == "uncertain" for value in st.session_state.atlas_quiz_results.values())
    wrong = sum(value == "wrong" for value in st.session_state.atlas_quiz_results.values())
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("صفحات مقروءة", len(st.session_state.atlas_seen))
    c2.metric("مؤشرة", len(st.session_state.atlas_marks))
    c3.metric("القبو الأحمر", len(st.session_state.atlas_red_zone))
    c4.metric("اختبارات الصور", known + uncertain + wrong)
    st.caption(f"عرفتها: {known} · متردد: {uncertain} · لم تعرفها: {wrong}")

    st.download_button("📥 تنزيل تقدم CSV", progress_csv(), f"Atlas_progress_{dt.date.today()}.csv", "text/csv", key="atlas_csv")
    st.download_button(
        "💾 تنزيل نسخة احتياطية JSON",
        json.dumps(backup_payload(), ensure_ascii=False, indent=2).encode("utf-8"),
        f"Atlas_backup_{dt.date.today()}.json",
        "application/json",
        key="atlas_json",
    )
    restore = st.file_uploader("♻️ استرجاع JSON من جهاز آخر", type=["json"], key="restore_atlas_json")
    if restore and st.button("استرجاع التقدم", key="restore_atlas_button"):
        st.session_state.restore_notice = restore_backup(restore.getvalue())
        st.rerun()
    if st.session_state.restore_notice:
        st.success(st.session_state.restore_notice)
        st.session_state.restore_notice = ""

    if st.session_state.atlas_page_meta:
        st.subheader("الصور الصعبة المسجلة")
        rows = []
        for pid in sorted(st.session_state.atlas_red_zone):
            meta = st.session_state.atlas_page_meta.get(pid, {})
            rows.append({"الأطلس": meta.get("source", ""), "PDF page": meta.get("pdf_page", ""), "المحور": meta.get("topic", "")})
        if rows:
            st.table(rows)
    st.warning("التقدم محفوظ داخل الجلسة فقط؛ استعمل JSON backup قبل تبديل الجهاز أو إعادة تشغيل التطبيق.")
    if st.button("🧹 تصفير تقدم الأطلس", key="reset_atlas_progress"):
        st.session_state.atlas_marks = set()
        st.session_state.atlas_red_zone = set()
        st.session_state.atlas_seen = set()
        st.session_state.atlas_page_meta = {}
        st.session_state.atlas_quiz_results = {}
        st.success("تم تصفير التقدم في الجلسة الحالية.")


st.title("🚀 Pediatric Hem/Onc Atlas Exam Hub 🩺")
st.caption(f"نسخة {APP_VERSION} · Atlas-only · لا يوجد MCQ")
st.info("🔒 خلي GitHub والتطبيق Private. لا تضع كتب الأطلس داخل المستودع؛ ارفعها من تبويب التحميل.")

tab_upload, tab_explorer, tab_exam, tab_progress = st.tabs(
    ["📥 تحميل الأطالس", "🔬 مستكشف الأطلس", "🧠 وضع الامتحان", "📈 التقدم"]
)

with tab_upload:
    loaded_sources = render_upload_tab()
with tab_explorer:
    render_explorer_tab(loaded_sources)
with tab_exam:
    render_exam_tab(loaded_sources)
with tab_progress:
    render_progress_tab()
