import streamlit as st
from pypdf import PdfReader
from docx import Document
import requests, re, json, io, html
from pathlib import Path
from datetime import datetime
from collections import Counter

st.set_page_config(page_title="AI Article Reviewer", page_icon="📝", layout="wide")

CSS = """
<style>
.main {max-width: 1250px; margin:auto;}
.hero {padding:22px 24px;border-radius:16px;background:linear-gradient(135deg,#172554,#1e3a8a);color:white;margin-bottom:18px;}
.hero h1{margin:0 0 6px 0}.card{padding:18px;border:1px solid #dbe3ef;border-radius:14px;background:#fff;margin-bottom:12px}
.small{color:#64748b;font-size:13px}.issue-critical{border-left:5px solid #b91c1c;padding:10px;background:#fef2f2;margin:7px 0}
.issue-major{border-left:5px solid #d97706;padding:10px;background:#fffbeb;margin:7px 0}.issue-minor{border-left:5px solid #2563eb;padding:10px;background:#eff6ff;margin:7px 0}
.issue-editorial{border-left:5px solid #64748b;padding:10px;background:#f8fafc;margin:7px 0;color:#0f172a}
.issue-critical,.issue-major,.issue-minor{color:#0f172a}
.finding-box{padding:14px 16px;background:#ffffff;color:#0f172a;border:1px solid #cbd5e1;border-left:5px solid #64748b;border-radius:8px;margin:8px 0;line-height:1.6}
.recommendation-box{padding:12px 15px;background:#f8fafc;color:#0f172a;border:1px solid #cbd5e1;border-radius:8px;margin:7px 0;line-height:1.6}
.solution-box{padding:12px 15px;background:#ffffff;color:#0f172a;border:1px solid #cbd5e1;border-left:4px solid #16a34a;border-radius:8px;margin:7px 0;line-height:1.6}
.summary-section{padding:16px 18px;background:#ffffff;color:#0f172a;border:1px solid #cbd5e1;border-radius:10px;margin:10px 0;line-height:1.65}
.summary-section h4{color:#0f172a;margin:0 0 8px 0}
.status-reviewed{color:#166534;font-weight:700}.status-pending{color:#92400e;font-weight:700}
.evidence-box{border-left:4px solid #64748b;padding:10px 12px;background:#f8fafc;color:#0f172a;border-radius:8px;margin:10px 0 0 0;font-size:14px;line-height:1.55}
.review-table-wrap{width:100%;overflow:visible}
.review-table{width:100%;border-collapse:collapse;table-layout:fixed;font-size:14px}
.review-table th,.review-table td{border:1px solid #cbd5e1;padding:8px 9px;vertical-align:top;white-space:normal;overflow-wrap:anywhere;word-break:break-word}
.review-table th{background:#eef2f7;font-weight:700}
.review-table .c-no{width:5%;text-align:center}
.review-table .c-type{width:18%}
.review-table .c-year{width:8%;text-align:center}
.review-table .c-match{width:11%;text-align:center}
.review-table .c-ref{width:58%}
.title-full{padding:14px 16px;border:1px solid #334155;border-radius:10px;background:#111827;font-size:18px;font-weight:700;line-height:1.55;white-space:normal;overflow-wrap:anywhere}
.scope-card{padding:14px 16px;border:1px solid #334155;border-radius:10px;background:#111827;margin:7px 0;line-height:1.5}
.sidebar-profile{padding:18px 14px 16px;border:1px solid rgba(148,163,184,.20);border-radius:18px;background:linear-gradient(180deg,rgba(30,41,59,.92),rgba(17,24,39,.92));box-shadow:0 10px 28px rgba(0,0,0,.20);margin-bottom:16px;text-align:center}
.sidebar-profile img{width:118px;height:118px;object-fit:cover;border-radius:50%;border:4px solid rgba(255,255,255,.95);box-shadow:0 7px 20px rgba(0,0,0,.30);display:block;margin:0 auto 12px}
.sidebar-profile .name{font-size:18px;font-weight:800;line-height:1.25;color:#f8fafc;margin:0}
.sidebar-profile .role{font-size:12px;line-height:1.45;color:#cbd5e1;margin:5px 0 0}
.sidebar-profile .line{height:1px;background:linear-gradient(90deg,transparent,rgba(148,163,184,.35),transparent);margin:14px 0 12px}
</style>
"""
st.markdown(CSS, unsafe_allow_html=True)

for k,v in {"article_text":"","filename":"","review":{},"history":[],"extract_meta":{},"scope_alternatives":[],"stage_reviews":{},"stage_text":{}}.items():
    if k not in st.session_state: st.session_state[k]=v

SECTIONS=["Judul","Abstrak","Pendahuluan","Metodologi","Hasil","Discussion","Conclusion","References"]


def extract_pdf(data):
    """Extract the complete PDF, page by page, without silently truncating text."""
    reader=PdfReader(io.BytesIO(data))
    pages=[]
    empty_pages=[]
    for idx, page in enumerate(reader.pages, start=1):
        txt=page.extract_text() or ""
        txt=txt.replace("\x00", "").strip()
        if not txt:
            empty_pages.append(idx)
        pages.append(txt)
    text="\n\n".join(f"[HALAMAN {i}]\n{txt}" if txt else f"[HALAMAN {i}]\n[TEKS TIDAK TERBACA]" for i,txt in enumerate(pages, start=1))
    return text, {"pages":len(pages), "empty_pages":empty_pages, "page_chars":[len(x) for x in pages]}


def extract_docx(data):
    """Extract the complete DOCX body and tables; no character slice is applied."""
    doc=Document(io.BytesIO(data))
    parts=[]
    for p in doc.paragraphs:
        if p.text.strip(): parts.append(p.text.strip())
    for table_idx, table in enumerate(doc.tables, start=1):
        parts.append(f"[TABEL {table_idx}]")
        for row in table.rows:
            parts.append(" | ".join(c.text.strip() for c in row.cells))
    text="\n".join(parts)
    return text, {"pages":None, "empty_pages":[], "tables":len(doc.tables), "paragraphs":len([p for p in doc.paragraphs if p.text.strip()])}


def normalize_space(s):
    return re.sub(r"\s+"," ",(s or "")).strip()


def _normalize_heading(line):
    line=re.sub(r"\s+"," ",(line or "").strip())
    line=re.sub(r"^\s*bab\s+[ivxlcdm0-9]+\s*[:.\-)]*\s*","",line,flags=re.I)
    line=re.sub(r"^\s*(?:\d+(?:\.\d+)*|[ivxlcdm]+)\s*[:.\-)]+\s*","",line,flags=re.I)
    return line.strip(" -:;|\t").strip()


def _heading_kind(line):
    raw=(line or "").strip()
    if not raw or len(raw)>120: return None
    h=_normalize_heading(raw).lower()
    aliases={
        "Abstrak":["abstrak","abstract"],
        "Pendahuluan":["pendahuluan","introduction","latar belakang"],
        "Metodologi":["metodologi","metodologi penelitian","metode penelitian","metode","methodology","methods","research method","research methodology"],
        "Hasil":["hasil","results","temuan","findings"],
        "Discussion":["pembahasan","discussion"],
        "Conclusion":["kesimpulan","simpulan","conclusion","penutup"],
        "References":["daftar pustaka","referensi","references","daftar rujukan","bibliography"],
    }
    for combined in ["hasil dan pembahasan","results and discussion","pembahasan dan hasil"]:
        if h==combined: return combined.title()
    for canonical,words in aliases.items():
        if h in words: return canonical
    return None


def _extract_title(text,first_section_pos=0):
    """Extract the complete article title, including wrapped title lines."""
    before=text[:first_section_pos] if first_section_pos else text[:8000]
    m=re.search(r"(?im)^\s*(?:judul|title)\s*[:\-]\s*(.+?)(?=\n\s*(?:abstrak|abstract|pendahuluan|introduction)\b|$)",before,re.S)
    if m:
        return normalize_space(m.group(1))

    lines=[x.strip() for x in before.splitlines() if x.strip()]
    if not lines: return ""

    # Stop at common author/institution metadata. Collect consecutive title-like lines.
    title=[]
    metadata_terms=("@","email","orcid","universitas","university","fakultas","program studi","department","doi","[halaman")
    stop_terms=("abstrak","abstract","kata kunci","keywords","pendahuluan","introduction")
    for line in lines[:25]:
        low=line.lower()
        if any(t in low for t in stop_terms): break
        if any(t in low for t in metadata_terms):
            if title: break
            continue
        if re.match(r"^(?:penulis|author|oleh)\s*[:\-]",low): break
        if re.match(r"^(?:[A-Z][a-z]+\s+){1,4}[A-Z][a-z]+(?:\s*,|\s*$)",line) and len(line.split())<=12 and not line.isupper():
            if title: break
        if 1<=len(line.split())<=45 and len(line)<=350:
            title.append(line)
            # If the title is a single long line, it is complete.
            if len(line.split())>=5 and len(title)==1 and len(line)>=35:
                # Continue only if next line also clearly looks like title text.
                continue
        elif title:
            break

    # Remove accidental trailing author line and join wrapped title lines.
    cleaned=[]
    for line in title:
        low=line.lower()
        if any(t in low for t in metadata_terms) or re.match(r"^(?:penulis|author|oleh)\b",low):
            break
        cleaned.append(line)
    return normalize_space(" ".join(cleaned[:4]))

def split_sections(text):
    """Split the complete article without character limits and with tolerant heading detection."""
    text=(text or "").replace("\r\n","\n").replace("\r","\n")
    lines=text.split("\n")
    starts=[]
    for i,line in enumerate(lines):
        kind=_heading_kind(line)
        if kind: starts.append((i,kind))

    out={k:"" for k in SECTIONS}
    # Canonical internal name for bibliography.
    out["References"]=""

    first_pos=sum(len(x)+1 for x in lines[:starts[0][0]]) if starts else 0
    out["Judul"]=_extract_title(text,first_pos)

    for idx,(line_no,kind) in enumerate(starts):
        next_line=starts[idx+1][0] if idx+1<len(starts) else len(lines)
        body="\n".join(lines[line_no+1:next_line]).strip()
        if kind in ("Hasil Dan Pembahasan","Results And Discussion","Pembahasan Dan Hasil"):
            # One source block can legitimately serve both result and discussion review.
            if not out["Hasil"]: out["Hasil"]=body
            if not out["Discussion"]: out["Discussion"]=body
        elif kind in out and not out[kind]:
            out[kind]=body
        elif kind=="References" and not out["References"]:
            out["References"]=body

    # Some documents put the bibliography after a visually styled heading that
    # PDF extraction does not preserve as a clean line. Use a conservative
    # fallback only when no reference heading was found.
    if not out["References"]:
        ref_match=re.search(r"(?is)(?:^|\n)\s*(?:daftar\s+pustaka|daftar\s+rujukan|referensi|references|bibliography)\s*[:.]?\s*\n(.*)$", text)
        if ref_match:
            out["References"]=ref_match.group(1).strip()

    # If a combined Results/Discussion heading is followed by Conclusion, keep
    # the same source text available to both review paths.
    if not out["Discussion"] and out["Hasil"]:
        # Do not fabricate discussion content when the document has a separate
        # discussion heading; this fallback is only for combined headings.
        combined_present=any(k in ["Hasil Dan Pembahasan","Results And Discussion","Pembahasan Dan Hasil"] for _,k in starts)
        if combined_present: out["Discussion"]=out["Hasil"]

    m=re.search(r"(?im)^\s*(?:judul|title)\s*[:\-]\s*(.+?)\s*$",text)
    if m: out["Judul"]=m.group(1).strip()
    return out


def extract_citations(text):
    patterns=[
        r"\(([A-Z][A-Za-zÀ-ÿ'’\-]+(?:\s+et\s+al\.)?(?:\s*&\s*[A-Z][A-Za-zÀ-ÿ'’\-]+)?[,\s]+(?:19|20)\d{2}[a-z]?)\)",
        r"\b[A-Z][A-Za-zÀ-ÿ'’\-]+(?:\s+et\s+al\.)?\s*\((?:19|20)\d{2}[a-z]?\)",
        r"\((?:19|20)\d{2}\)"
    ]
    found=[]
    for p in patterns: found.extend(re.findall(p,text))
    return found


def extract_years(text):
    return [int(y) for y in re.findall(r"\b(19\d{2}|20\d{2})\b",text)]


def detect_research_type(text):
    low=text.lower()
    quantitative=sum(low.count(x) for x in ["kuantitatif","quantitative","regresi","anova","uji t","uji f","korelasi","sampel"])
    qualitative=sum(low.count(x) for x in ["kualitatif","qualitative","wawancara","observasi partisipan","informan","analisis tematik","fenomenologi","etnografi"])
    mixed=sum(low.count(x) for x in ["mixed method","mixed methods","metode campuran","mixed-method"])
    if mixed>=1 or (quantitative>=2 and qualitative>=2): return "Mixed Methods"
    if quantitative>qualitative and quantitative>=2: return "Kuantitatif"
    if qualitative>quantitative and qualitative>=2: return "Kualitatif"
    if any(x in low for x in ["pengembangan","research and development","r&d","design-based"]): return "Pengembangan / R&D"
    return "Belum dapat ditentukan secara meyakinkan"


def contains_any(text,terms):
    low=text.lower(); return any(t.lower() in low for t in terms)


def issue(severity,section,problem,evidence="",action=""):
    return {"severity":severity,"section":section,"issue":problem,"evidence":evidence,"action":action}


def sentence_list(text):
    """Split text into evidence-ready sentences while preserving the manuscript wording."""
    cleaned=re.sub(r"\s+", " ", text or "").strip()
    if not cleaned: return []
    return [x.strip() for x in re.split(r"(?<=[.!?])\s+", cleaned) if x.strip()]


def _complete_quote(text, max_chars=520):
    """Return complete sentence(s), never a mid-sentence fragment, for evidence display."""
    sentences=sentence_list(text)
    if not sentences: return ""
    chosen=[]
    total=0
    for sent in sentences:
        extra=len(sent) + (1 if chosen else 0)
        if chosen and total + extra > max_chars:
            break
        if not chosen and len(sent) > max_chars:
            # Keep the full sentence when it is long; the UI can wrap it safely.
            return sent
        chosen.append(sent)
        total += extra
    return " ".join(chosen)


def find_evidence(text, terms, max_chars=520):
    """Return complete, verbatim sentence(s) containing an indicator."""
    sentences=sentence_list(text)
    low_terms=[t.lower() for t in terms if t and t.strip()]
    for idx,sent in enumerate(sentences):
        low=sent.lower()
        if any(t in low for t in low_terms):
            block=sent
            if idx+1 < len(sentences) and len(block) < max_chars*0.55:
                candidate=block + " " + sentences[idx+1]
                if len(candidate) <= max_chars:
                    block=candidate
            return _complete_quote(block,max_chars)
    return ""


def evidence_location(text, quote):
    """Give a traceable textual location without inventing page numbers."""
    q=normalize_space(quote)
    if not q:
        return "lokasi kutipan belum dapat ditentukan"
    raw_lines=(text or "").replace("\r\n","\n").replace("\r","\n").split("\n")
    norm_q=q.lower()
    page=None
    sentence_no=0
    for line in raw_lines:
        marker=re.search(r"\[HALAMAN\s+(\d+)\]",line,re.I)
        if marker:
            page=marker.group(1)
        if line.strip() and not re.match(r"^\[HALAMAN\s+\d+\]$",line.strip(),re.I):
            for sent in sentence_list(line):
                sentence_no += 1
                if norm_q[:80] in sent.lower() or sent.lower() in norm_q:
                    return f"kalimat sekitar ke-{sentence_no}" + (f", halaman {page}" if page else "")
    return (f"sekitar kalimat ke-{sentence_no}" if sentence_no else "lokasi tekstual") + (f", halaman {page}" if page else "")


def _evidence_label(label):
    return (label or "unspecified").strip().lower()


def evidence_for_missing(text, terms, label, section=""):
    """Produce a defensible negative textual finding; never infer absence from silence."""
    words=len((text or "").split())
    sec=section or "bagian yang dipilih"
    lab=_evidence_label(label)
    indicators=", ".join(terms[:8]) if terms else "indikator tekstual yang relevan"
    specific={
        "tujuan":"Penelusuran difokuskan pada kalimat yang menyatakan tujuan, sasaran, pertanyaan penelitian, atau rumusan capaian studi.",
        "masalah":"Penelusuran difokuskan pada uraian fenomena, persoalan ilmiah, kondisi empiris, atau alasan akademik yang melandasi penelitian.",
        "state of the art":"Penelusuran difokuskan pada sintesis penelitian terdahulu yang menunjukkan posisi pengetahuan yang telah tersedia dan bukan sekadar daftar sitasi.",
        "gap":"Penelusuran difokuskan pada keterbatasan studi terdahulu, pertanyaan yang belum terjawab, perbedaan temuan, atau ruang pengetahuan yang secara logis membuka kebutuhan penelitian baru.",
        "desain/pendekatan":"Penelusuran difokuskan pada pernyataan desain/pendekatan dan uraian yang menunjukkan bagaimana desain tersebut digunakan.",
        "subjek/objek/sumber data":"Penelusuran difokuskan pada siapa atau apa yang diteliti, unit analisis, populasi, sampel, informan, responden, atau sumber data beserta keterangan pemilihannya bila relevan.",
        "pengumpulan data":"Penelusuran difokuskan pada instrumen, teknik, prosedur, waktu, dan tahapan memperoleh data.",
        "teknik pengumpulan data":"Penelusuran difokuskan pada instrumen, teknik, prosedur, waktu, dan tahapan memperoleh data.",
        "analisis data":"Penelusuran difokuskan pada teknik, tahapan, kriteria, dan keluaran analisis yang digunakan untuk menjawab pertanyaan penelitian.",
        "teknik analisis data":"Penelusuran difokuskan pada teknik, tahapan, kriteria, dan keluaran analisis yang digunakan untuk menjawab pertanyaan penelitian.",
        "validitas/reliabilitas atau trustworthiness":"Penelusuran difokuskan pada prosedur yang menjamin kualitas instrumen atau keabsahan/keterpercayaan data sesuai desain penelitian.",
        "etika penelitian":"Penelusuran difokuskan pada persetujuan partisipan, kerahasiaan, perlindungan data/subjek, persetujuan etik, dan prinsip etika lain yang relevan.",
        "hasil":"Penelusuran difokuskan pada temuan empiris yang menjawab tujuan penelitian dan memiliki dasar data yang dapat ditelusuri.",
        "perbandingan literatur":"Penelusuran difokuskan pada hubungan antara temuan penelitian ini dengan studi terdahulu, termasuk persamaan, perbedaan, dan penjelasan atasnya.",
        "kontribusi":"Penelusuran difokuskan pada sumbangan penelitian terhadap teori, metode, praktik, kebijakan, atau pengembangan pengetahuan.",
        "implikasi/kontribusi":"Penelusuran difokuskan pada implikasi atau sumbangan penelitian yang benar-benar dapat diturunkan dari temuan.",
    }
    rationale=specific.get(lab, f"Penelusuran difokuskan pada indikator yang diperlukan untuk menilai komponen {label}.")
    if lab in {"gap","state of the art","masalah","tujuan","kontribusi","implikasi/kontribusi"}:
        closing="Hasil ini hanya menyatakan bahwa bukti eksplisit belum ditemukan pada bagian yang diperiksa. Reviewer tidak menyimpulkan bahwa unsur tersebut sama sekali tidak ada dalam penelitian sebelum bagian lain ditelusuri."
    elif lab in {"desain/pendekatan","subjek/objek/sumber data","pengumpulan data","teknik pengumpulan data","analisis data","teknik analisis data","validitas/reliabilitas atau trustworthiness","etika penelitian"}:
        closing="Karena unsur ini menentukan keterlacakan dan/atau replikasi penelitian, penulis perlu memperjelas informasi yang belum dapat diverifikasi dari naskah. Reviewer tidak boleh mengisi kekosongan tersebut dengan asumsi tentang prosedur yang sebenarnya dilakukan."
    else:
        closing="Kesimpulan dibatasi pada keterbacaan tekstual pada bagian yang diperiksa."
    return (
        f"Status bukti: BELUM TERDETEKSI SECARA EKSPLISIT. {rationale} "
        f"Pada {sec}, sekitar {words:,} kata ditelusuri menggunakan indikator seperti {indicators}. "
        f"Tidak ditemukan pernyataan yang cukup jelas untuk memverifikasi komponen {label} pada bagian tersebut. {closing}"
    )


def evidence_for_present(text, terms, label, section="", max_chars=520):
    """Produce a positive evidence note with verbatim quote, location, and interpretation."""
    quote=find_evidence(text, terms, max_chars=max_chars)
    if not quote:
        return evidence_for_missing(text, terms, label, section)
    lab=_evidence_label(label)
    interpretations={
        "tujuan":"Kutipan ini dapat digunakan untuk memverifikasi keberadaan tujuan penelitian; kecukupan tujuan tetap dinilai dari keterhubungannya dengan pertanyaan, metode, hasil, dan kesimpulan.",
        "masalah":"Kutipan ini menunjukkan dasar tekstual persoalan yang dibangun penulis; reviewer selanjutnya menilai apakah persoalan tersebut memiliki signifikansi ilmiah dan didukung literatur/empiri.",
        "state of the art":"Kutipan ini menunjukkan upaya menempatkan penelitian melalui literatur terdahulu; kualitas state of the art ditentukan oleh sintesis, relevansi, dan posisi pengetahuan yang dibangun.",
        "gap":"Kutipan ini memuat indikator kesenjangan yang dapat ditelusuri; reviewer perlu menilai apakah gap tersebut benar-benar diturunkan dari literatur dan relevan dengan tujuan penelitian.",
        "desain/pendekatan":"Kutipan ini menjadi dasar untuk mengidentifikasi desain atau pendekatan yang dinyatakan penulis; kecukupannya dinilai dari konsistensi dengan data dan analisis.",
        "subjek/objek/sumber data":"Kutipan ini memberikan informasi tentang sumber data atau unit yang diteliti; kelengkapannya perlu dinilai dari jumlah, karakteristik, lokasi, dan kriteria pemilihan bila relevan.",
        "pengumpulan data":"Kutipan ini menunjukkan teknik/prosedur memperoleh data; reviewer perlu memeriksa apakah langkahnya cukup rinci untuk ditelusuri.",
        "teknik pengumpulan data":"Kutipan ini menunjukkan teknik/prosedur memperoleh data; reviewer perlu memeriksa apakah langkahnya cukup rinci untuk ditelusuri.",
        "analisis data":"Kutipan ini menunjukkan teknik atau prosedur analisis; kecukupannya dinilai dari kesesuaian dengan desain, jenis data, dan pertanyaan penelitian.",
        "teknik analisis data":"Kutipan ini menunjukkan teknik atau prosedur analisis; kecukupannya dinilai dari kesesuaian dengan desain, jenis data, dan pertanyaan penelitian.",
        "validitas/reliabilitas atau trustworthiness":"Kutipan ini menunjukkan adanya prosedur penjaminan kualitas data/instrumen; reviewer perlu memeriksa prosedur dan hasil penerapannya.",
        "etika penelitian":"Kutipan ini menjadi dasar untuk menilai pernyataan etika; kecukupannya bergantung pada karakter subjek, jenis data, dan risiko penelitian.",
        "hasil":"Kutipan ini menunjukkan temuan/evidensi hasil; reviewer perlu memastikan bahwa temuan tersebut benar-benar didukung data yang dilaporkan.",
        "perbandingan literatur":"Kutipan ini menunjukkan upaya membandingkan temuan dengan studi terdahulu; kualitas pembahasan ditentukan oleh kedalaman interpretasi, bukan sekadar adanya sitasi.",
        "kontribusi":"Kutipan ini menjadi dasar untuk menilai klaim kontribusi; klaim harus proporsional dengan bukti dan tidak melampaui temuan penelitian.",
        "implikasi/kontribusi":"Kutipan ini menjadi dasar untuk menilai implikasi atau kontribusi; klaim harus proporsional dengan bukti penelitian.",
    }
    interpretation=interpretations.get(lab, f"Kutipan ini merupakan dasar tekstual untuk menilai komponen {label}.")
    loc=evidence_location(text,quote)
    return f"Status bukti: TERKONFIRMASI. Lokasi tekstual: {section or 'bagian yang dipilih'}, {loc}. Cuplikan naskah verbatim: “{quote}” {interpretation}"

def grounded_finding(text, label, terms, severity, problem, action):
    """Create a reviewer finding that always carries manuscript-grounded evidence."""
    evidence=find_evidence(text, terms)
    if evidence and any(t.lower() in evidence.lower() for t in terms):
        ev=f'Cuplikan naskah: “{evidence}”'
    else:
        ev=evidence_for_present(text, terms, label, "") if evidence else evidence_for_missing(text, terms, label, "")
    return issue(severity, label, problem, ev, action)


def recommendation_for_methodology(text, research_type=None):
    """Create methodology advice that is specific to the detected design and missing components."""
    rt=research_type or detect_research_type(text)
    checks=[
        ("desain/pendekatan",["pendekatan","desain","design","method","metode"]),
        ("subjek/objek/sumber data",["subjek","objek","informan","responden","sampel","populasi","sumber data"]),
        ("teknik pengumpulan data",["wawancara","observasi","dokumentasi","kuesioner","pengumpulan data","data collection"]),
        ("teknik analisis data",["analisis data","data analysis","coding","analisis tematik","regresi","anova","uji t","uji f"]),
        ("validitas/reliabilitas atau trustworthiness",["validitas","reliabilitas","reliability","validity","trustworthiness","credibility","triangulasi"]),
        ("etika penelitian",["etika penelitian","ethical","informed consent","persetujuan etik","komite etik"]),
    ]
    missing=[label for label,terms in checks if not contains_any(text,terms)]
    rec=[]
    if rt=="Kuantitatif":
        templates={
            "desain/pendekatan":"Tuliskan secara eksplisit desain kuantitatif yang digunakan (misalnya survei, korelasional, eksperimen, atau ex post facto) serta alasan pemilihannya sesuai tujuan penelitian.",
            "subjek/objek/sumber data":"Sebutkan populasi, unit analisis, teknik sampling, jumlah sampel, serta kriteria inklusi/eksklusi jika digunakan. Jelaskan juga sumber data primer atau sekunder.",
            "teknik pengumpulan data":"Jelaskan instrumen, indikator/variabel yang diukur, prosedur pengumpulan data, dan waktu pelaksanaannya agar proses dapat ditelusuri.",
            "teknik analisis data":"Sebutkan teknik statistik yang digunakan, urutan analisis, asumsi/kriteria pengujian yang relevan, dan alasan teknik tersebut sesuai dengan jenis data serta pertanyaan penelitian.",
            "validitas/reliabilitas atau trustworthiness":"Laporkan bukti validitas dan reliabilitas instrumen yang benar-benar dilakukan, termasuk teknik dan hasil utamanya.",
            "etika penelitian":"Jelaskan persetujuan partisipan, kerahasiaan data, persetujuan etik bila diwajibkan, dan perlindungan terhadap partisipan sesuai karakter penelitian."
        }
    elif rt=="Kualitatif":
        templates={
            "desain/pendekatan":"Nyatakan pendekatan/desain kualitatif yang digunakan (misalnya fenomenologi, studi kasus, etnografi, atau deskriptif) dan jelaskan alasan kesesuaiannya dengan pertanyaan penelitian.",
            "subjek/objek/sumber data":"Jelaskan siapa informan/subjek penelitian, kriteria pemilihannya, jumlah atau kecukupan informan, lokasi, serta sumber data yang digunakan.",
            "teknik pengumpulan data":"Uraikan prosedur wawancara, observasi, dokumentasi, atau teknik lain secara operasional, termasuk waktu dan cara pencatatan data.",
            "teknik analisis data":"Jelaskan tahap analisis, seperti transkripsi, coding, kategorisasi/tema, interpretasi, dan cara keputusan analitis dibuat.",
            "validitas/reliabilitas atau trustworthiness":"Jelaskan strategi credibility, dependability, confirmability, transferability, triangulasi, member checking, atau strategi lain yang benar-benar digunakan.",
            "etika penelitian":"Jelaskan persetujuan informan, kerahasiaan identitas, penggunaan data, dan persetujuan etik bila diwajibkan."
        }
    elif rt=="Mixed Methods":
        templates={
            "desain/pendekatan":"Nyatakan desain mixed methods yang digunakan dan jelaskan urutan, prioritas, serta hubungan fase kuantitatif dan kualitatif.",
            "subjek/objek/sumber data":"Jelaskan populasi/sampel dan informan pada masing-masing fase serta dasar pemilihannya.",
            "teknik pengumpulan data":"Jelaskan instrumen dan prosedur pengumpulan data untuk kedua fase secara terpisah.",
            "teknik analisis data":"Jelaskan analisis kuantitatif dan kualitatif masing-masing, lalu tunjukkan pada tahap mana kedua hasil diintegrasikan.",
            "validitas/reliabilitas atau trustworthiness":"Laporkan validitas/reliabilitas untuk komponen kuantitatif dan strategi trustworthiness untuk komponen kualitatif sesuai yang benar-benar dilakukan.",
            "etika penelitian":"Jelaskan persetujuan dan perlindungan partisipan pada seluruh fase penelitian."
        }
    elif "Pengembangan" in rt:
        templates={
            "desain/pendekatan":"Sebutkan model pengembangan yang digunakan dan uraikan setiap tahap yang benar-benar dilaksanakan.",
            "subjek/objek/sumber data":"Jelaskan subjek uji coba, validator, jumlahnya, kriteria pemilihan, serta karakteristik produk yang dikembangkan.",
            "teknik pengumpulan data":"Jelaskan instrumen validasi/uji coba, prosedur pelaksanaan, dan cara data kelayakan atau kepraktisan diperoleh.",
            "teknik analisis data":"Jelaskan rumus/teknik analisis, kriteria kelayakan, dan dasar pengambilan keputusan terhadap produk.",
            "validitas/reliabilitas atau trustworthiness":"Jelaskan validasi ahli atau pengujian instrumen yang benar-benar dilakukan serta hasilnya.",
            "etika penelitian":"Jelaskan persetujuan dan perlindungan peserta uji coba jika penelitian melibatkan manusia."
        }
    else:
        templates={label:f"Jelaskan {label} secara operasional sesuai desain penelitian yang benar-benar dilakukan, termasuk prosedur dan dasar pengambilannya." for label,_ in checks}

    if missing:
        for label in missing:
            rec.append(f"{len(rec)+1}. {templates[label]}")
        rec.append(f"{len(rec)+1}. Setelah komponen di atas dilengkapi, periksa kembali keselarasan metodologi dengan tujuan, jenis data, hasil, dan kesimpulan. Jangan menambahkan prosedur yang tidak benar-benar dilakukan.")
    else:
        rec.append("1. Komponen metodologi utama telah terdeteksi secara tekstual. Tahap berikutnya adalah memeriksa kecukupan detail dan konsistensi prosedur dengan data, hasil, serta kesimpulan.")
        rec.append("2. Pastikan setiap prosedur yang dilaporkan dapat ditelusuri ke data yang benar-benar dianalisis dan tidak ada prosedur yang ditambahkan hanya untuk memenuhi format jurnal.")
    return " ".join(rec)


def section_checks(sec):
    findings=[]; strengths=[]; rec=[]
    for name,text in sec.items():
        if name=="References": continue
        w=len(text.split())
        if text: strengths.append((name,w))
        else: findings.append(issue("MAJOR",name,f"Bagian {name} belum berhasil diekstraksi.","Tidak ada teks yang terdeteksi pada heading bagian ini.","Periksa heading dokumen atau masukkan bagian secara manual."))
    return findings,strengths


def demo_review(text,meta):
    sec=split_sections(text); words=len(text.split()); citations=extract_citations(text); years=extract_years(text)
    issues=[]; roadmap=[]
    rtype=detect_research_type(text)
    title=sec["Judul"]; intro=sec["Pendahuluan"]; method=sec["Metodologi"]; result=sec["Hasil"]; disc=sec["Discussion"]; concl=sec["Conclusion"]; abstract=sec["Abstrak"]

    # Structural checks: absence, not just word count.
    if not abstract: issues.append(issue("MAJOR","Abstrak","Abstrak tidak berhasil diekstraksi.","Tidak ditemukan heading Abstrak/Abstract.","Periksa heading atau gunakan Review Bertahap."))
    else:
        for label,terms in [("tujuan",["tujuan","bertujuan","aim","objective"]),("metode",["metode","method","pendekatan","approach"]),("hasil",["hasil","temuan","result","finding"]),("kontribusi",["kontribusi","implikasi","contribution","implication"])]:
            if not contains_any(abstract,terms): issues.append(issue("MINOR","Abstrak",f"Komponen {label} belum terdeteksi secara tekstual.","Pemeriksaan berbasis kata kunci; verifikasi makna kalimat secara akademik.",f"Pastikan abstrak menyatakan {label} secara eksplisit dan ringkas."))
    if not intro: issues.append(issue("MAJOR","Pendahuluan","Pendahuluan tidak berhasil diekstraksi.","Tidak ditemukan heading Pendahuluan/Introduction.","Periksa heading dokumen."))
    else:
        for label,terms in [("masalah",["masalah","problem","fenomena","isu","permasalahan"]),("state of the art",["penelitian terdahulu","studi terdahulu","literatur","state of the art","previous studies"]),("gap",["research gap","research gap","kesenjangan","belum","masih terbatas"]),("tujuan",["tujuan penelitian","bertujuan","aim","objective"])]:
            if not contains_any(intro,terms): issues.append(issue("MAJOR" if label in ["masalah","gap","tujuan"] else "MINOR","Pendahuluan",f"Elemen {label} belum terdeteksi secara tekstual.","Pemeriksaan berbasis indikator tekstual.",f"Bangun alur masalah → state of the art → gap → tujuan → kontribusi."))
    if not method: issues.append(issue("CRITICAL","Metodologi","Metodologi tidak berhasil diekstraksi sehingga kelayakan desain dan replikasi belum dapat dinilai.","Tidak ditemukan heading Metodologi/Metode/Methodology.","Periksa heading dan jelaskan desain, sumber data, sampel/informan, instrumen, prosedur, analisis, validitas/reliabilitas atau trustworthiness, serta etika."))
    else:
        checks=[("desain/pendekatan",["pendekatan","desain","design","method","metode"]),("sumber data atau subjek",["subjek","objek","informan","responden","sampel","populasi","sumber data"]),("pengumpulan data",["wawancara","observasi","dokumentasi","kuesioner","pengumpulan data","data collection"]),("analisis data",["analisis data","data analysis","coding","analisis tematik","regresi","anova","uji t"]),("validitas/reliabilitas atau trustworthiness",["validitas","reliabilitas","reliability","validity","trustworthiness","credibility"]),("etika",["etik","etika penelitian","ethical","informed consent","persetujuan etik"]) ]
        for label,terms in checks:
            if not contains_any(method,terms): issues.append(issue("MAJOR" if label in ["desain/pendekatan","analisis data"] else "MINOR","Metodologi",f"Komponen {label} belum terdeteksi secara tekstual.","Tidak ditemukan indikator yang cukup untuk komponen tersebut.",f"Tambahkan penjelasan {label} sesuai desain penelitian."))
    if not result: issues.append(issue("MAJOR","Hasil","Bagian hasil tidak berhasil diekstraksi.","Tidak ditemukan heading Hasil/Results/Findings.","Pastikan temuan utama disajikan dengan evidensi yang dapat ditelusuri."))
    else:
        if not contains_any(result,["tabel","table","gambar","figure","menunjukkan","temuan","hasil penelitian"]): issues.append(issue("MINOR","Hasil","Evidensi hasil belum terdeteksi secara tekstual dengan kuat.","Tidak ditemukan indikator tabel/gambar/temuan/hasil.","Pastikan temuan utama didukung data, tabel, gambar, kutipan, atau ukuran efek sesuai desain."))
    if not disc: issues.append(issue("MAJOR","Discussion","Pembahasan tidak berhasil diekstraksi.","Tidak ditemukan heading Pembahasan/Discussion.","Bangun interpretasi temuan dan bandingkan dengan penelitian terdahulu."))
    else:
        if not contains_any(disc,["penelitian terdahulu","studi terdahulu","sejalan","berbeda","dibandingkan","literatur","previous studies"]): issues.append(issue("MAJOR","Discussion","Perbandingan hasil dengan penelitian terdahulu belum terdeteksi secara tekstual.","Indikator komparasi literatur tidak ditemukan.","Jelaskan persamaan/perbedaan hasil dengan studi terdahulu dan alasan ilmiahnya."))
        if not contains_any(disc,["implikasi","kontribusi","contribution","implication","makna","berarti"]): issues.append(issue("MINOR","Discussion","Implikasi atau kontribusi temuan belum tampak kuat secara tekstual.","Indikator implikasi/kontribusi tidak ditemukan.","Jelaskan kontribusi teoritis, metodologis, praktis, atau kebijakan sesuai penelitian."))
    if not concl: issues.append(issue("MINOR","Conclusion","Kesimpulan tidak berhasil diekstraksi.","Tidak ditemukan heading Kesimpulan/Conclusion.","Pastikan kesimpulan menjawab tujuan dan berbasis hasil."))
    elif len(concl.split())<60: issues.append(issue("MINOR","Conclusion","Kesimpulan relatif singkat.",f"Sekitar {len(concl.split())} kata.","Pastikan seluruh tujuan/research question terjawab tanpa menambah klaim baru."))

    # Internal consistency indicators.
    if title and intro and method:
        title_words=set(re.findall(r"[a-zA-ZÀ-ÿ]{5,}",title.lower()))
        intro_words=set(re.findall(r"[a-zA-ZÀ-ÿ]{5,}",intro.lower()))
        overlap=len(title_words & intro_words)/max(1,len(title_words))
        if overlap<0.25: issues.append(issue("MAJOR","Konsistensi","Fokus judul memiliki sedikit keterhubungan leksikal dengan pendahuluan.",f"Tumpang tindih kata kunci sekitar {overlap:.0%}; ini bukan bukti definitif ketidaksesuaian.","Periksa secara substantif apakah objek, variabel/fokus, dan tujuan konsisten dengan judul."))

    # Reference audit.
    ref_section=sec.get("References","")
    ref_lines=[x.strip() for x in ref_section.splitlines() if x.strip()]
    if not ref_section: issues.append(issue("MAJOR","References","Daftar pustaka tidak berhasil diekstraksi.","Tidak ditemukan heading References/Daftar Pustaka.","Periksa struktur dokumen dan lakukan audit referensi."))
    elif len(ref_lines)<5: issues.append(issue("MAJOR","References","Daftar pustaka yang terdeteksi sangat sedikit.",f"Hanya sekitar {len(ref_lines)} baris entri terdeteksi.","Verifikasi kelengkapan referensi dan kecocokan sitasi dengan daftar pustaka."))
    if years:
        recent=sum(1 for y in years if y>=2021)
        old=sum(1 for y in years if y<2016)
        if recent==0: issues.append(issue("MINOR","References","Belum terdeteksi referensi 2021–sekarang dari teks yang dibaca.","Pemeriksaan tahun berbasis teks; tahun dapat berasal dari metadata lain.","Perbarui sebagian literatur dengan sumber primer mutakhir yang relevan."))
        if old>recent*2 and len(years)>5: issues.append(issue("MINOR","References","Proporsi tahun referensi lama tampak dominan.",f"Tahun lama (<2016): {old}; tahun 2021+: {recent}.","Pertahankan sumber klasik yang fundamental, tetapi tambahkan studi primer mutakhir."))

    issues=ground_full_review_issues(issues, sec)

    # Score from measurable domains; not a fake exact judgment.
    weights={"Title & Abstract":10,"Introduction & Gap":15,"Literature & Framework":15,"Methodology":20,"Results":10,"Discussion":15,"Conclusion":5,"References":5,"Writing":5}
    deductions={k:0 for k in weights}
    for it in issues:
        sev=it["severity"]; d={"CRITICAL":12,"MAJOR":6,"MINOR":2,"EDITORIAL":1}.get(sev,1)
        s=it["section"]
        key=("Methodology" if s=="Metodologi" else "Discussion" if s=="Discussion" else "Results" if s=="Hasil" else "Conclusion" if s=="Conclusion" else "References" if s=="References" else "Introduction & Gap" if s=="Pendahuluan" else "Title & Abstract" if s in ["Judul","Abstrak"] else "Literature & Framework")
        deductions[key]+=d
    base=100
    score=max(0,round(base-sum(min(v,18) for v in deductions.values())))
    if any(i["severity"]=="CRITICAL" for i in issues): rec="REJECT AND RESUBMIT"
    elif score>=90: rec="ACCEPT"
    elif score>=80: rec="MINOR REVISION"
    elif score>=65: rec="MAJOR REVISION"
    else: rec="REJECT AND RESUBMIT"
    priorities=["Substansi","Metodologi","Gap/Novelty","Konsistensi","Referensi"]
    for p in priorities: roadmap.append(f"Periksa dan perbaiki {p.lower()} berdasarkan temuan reviewer; gunakan evidensi dari artikel, bukan asumsi.")
    roadmap += ["Sinkronkan Judul → Tujuan → Research Question → Metode → Data → Hasil → Discussion → Conclusion → Abstract.","Audit sitasi terhadap daftar pustaka, termasuk sumber primer, relevansi, kemutakhiran, dan konsistensi gaya.","Lakukan penyuntingan bahasa dan format setelah substansi selesai."]

    gap="Belum layak menyatakan research gap hanya dari kata kunci. Indikator yang terdeteksi harus diverifikasi dengan sintesis penelitian terdahulu dan state of the art." if not intro else ("Pendahuluan memuat indikator gap, tetapi validitas gap tetap memerlukan sintesis komparatif terhadap penelitian terdahulu." if contains_any(intro,["gap","kesenjangan","belum","masih terbatas"]) else "Indikator research gap belum terdeteksi secara jelas pada pendahuluan.")
    novelty="Novelty belum dapat dibuktikan secara valid tanpa membandingkan kontribusi artikel dengan literatur primer yang relevan." 
    summary=f"Artikel sekitar {words:,} kata. Pemeriksaan offline mendeteksi tipe penelitian sementara: {rtype}. Sistem menemukan {len(issues)} isu terstruktur dan {len(citations)} indikator sitasi; hasil ini merupakan diagnosis berbasis teks, bukan pengganti penilaian reviewer manusia."
    return {"summary":summary,"sections":sec,"gap":gap,"novelty":novelty,"method":f"Jenis penelitian terindikasi: {rtype}. Metodologi diperiksa berdasarkan desain, sumber/subjek, pengumpulan data, analisis, validitas/reliabilitas atau trustworthiness, dan etika.","consistency":"Periksa silang Judul → Tujuan → Research Question → Metode → Data → Hasil → Discussion → Conclusion → Abstract. Sistem juga memeriksa indikator keterhubungan fokus judul dengan pendahuluan.","issues":issues,"score":score,"recommendation":rec,"roadmap":roadmap,"metrics":{"words":words,"citations":len(citations),"reference_lines":len(ref_lines),"research_type":rtype,"section_words":{k:len(v.split()) for k,v in sec.items() if k!="References"},"score_domains":{k:max(0,weights[k]-deductions[k]) for k in weights}}}


def ground_full_review_issues(issues, sec):
    """Replace generic heuristic evidence with grounded manuscript excerpts where possible."""
    term_map={
        "abstrak": ["tujuan","bertujuan","metode","method","hasil","temuan","result","kontribusi","implikasi"],
        "pendahuluan": ["masalah","problem","fenomena","isu","penelitian terdahulu","studi terdahulu","literatur","state of the art","research gap","kesenjangan","belum","masih terbatas","tujuan penelitian","bertujuan"],
        "metodologi": ["pendekatan","desain","design","method","metode","subjek","objek","informan","responden","sampel","populasi","sumber data","wawancara","observasi","dokumentasi","kuesioner","analisis data","data analysis","regresi","anova","uji t","validitas","reliabilitas","trustworthiness","etika penelitian","informed consent"],
        "hasil": ["tabel","table","gambar","figure","menunjukkan","temuan","hasil penelitian"],
        "discussion": ["penelitian terdahulu","studi terdahulu","sejalan","berbeda","dibandingkan","literatur","implikasi","kontribusi","contribution","implication"],
        "conclusion": ["berdasarkan","hasil","menyimpulkan","kesimpulan","simpulan"],
        "references": ["2021","2022","2023","2024","2025","2026","doi","jurnal","penerbit","isbn"],
        "judul": []
    }
    label_terms={
        "masalah":["masalah","problem","fenomena","isu"],
        "state of the art":["penelitian terdahulu","studi terdahulu","literatur","state of the art"],
        "gap":["research gap","kesenjangan","belum","masih terbatas"],
        "tujuan":["tujuan penelitian","bertujuan","aim","objective"],
        "desain":["pendekatan","desain","design","method","metode"],
        "subjek":["subjek","objek","informan","responden","sampel","populasi","sumber data"],
        "pengumpulan":["wawancara","observasi","dokumentasi","kuesioner","pengumpulan data","data collection"],
        "analisis":["analisis data","data analysis","coding","regresi","anova","uji t","uji f"],
        "validitas":["validitas","reliabilitas","reliability","validity","trustworthiness","credibility","triangulasi"],
        "etika":["etika penelitian","ethical","informed consent","persetujuan etik","komite etik"],
        "hasil":["tabel","table","gambar","figure","temuan","menunjukkan","hasil penelitian"],
        "perbandingan":["penelitian terdahulu","studi terdahulu","sejalan","berbeda","dibandingkan","literatur"],
        "implikasi":["implikasi","kontribusi","contribution","implication"]
    }
    for x in issues or []:
        if not isinstance(x,dict): continue
        sec_name=x.get("section","")
        text=sec.get(sec_name,"")
        if not text:
            continue
        problem=(x.get("issue","") or "").lower()
        chosen=[]; label=sec_name
        for key,terms in label_terms.items():
            if key in problem:
                chosen=terms; label=key; break
        if not chosen: chosen=term_map.get(sec_name.lower(),[])
        ev=find_evidence(text,chosen)
        if ev and any(t.lower() in ev.lower() for t in chosen):
            x["evidence"]=f'Cuplikan naskah: “{ev}”'
        elif not x.get("evidence") or "pemeriksaan berbasis" in str(x.get("evidence","")).lower() or "indikator" in str(x.get("evidence","")).lower():
            x["evidence"]=evidence_for_missing(text,chosen,label,sec_name)
    return issues


def professional_recommendation(section, label, text, research_type=""):
    """Create an actionable reviewer recommendation tailored to the finding type."""
    sec=section or "bagian yang diperiksa"
    lab=(label or "").lower()
    rt=research_type or detect_research_type(text)
    if "tujuan" in lab:
        return "Rumuskan tujuan penelitian dalam satu pernyataan yang eksplisit dan dapat diuji melalui metode serta hasil yang dilaporkan. Pastikan kata kerja tujuan menunjukkan apa yang dianalisis, dijelaskan, dibandingkan, atau dievaluasi, lalu cek kesesuaiannya dengan judul dan kesimpulan."
    if "masalah" in lab:
        return "Perjelas persoalan ilmiah dengan urutan fenomena atau bukti empiris → konsekuensi/masalah → keterbatasan pengetahuan saat ini. Hindari menjadikan pernyataan umum sebagai masalah penelitian tanpa dukungan data atau literatur yang relevan."
    if "state of the art" in lab:
        return "Sintesis penelitian terdahulu perlu menunjukkan apa yang telah diketahui, pendekatan yang telah digunakan, dan batas pengetahuan yang masih tersisa. Susun literatur secara argumentatif, bukan sebagai daftar ringkasan sumber satu per satu."
    if "gap" in lab:
        return "Nyatakan gap secara spesifik setelah membandingkan temuan atau pendekatan studi terdahulu. Tunjukkan apa yang belum terjawab, mengapa kekosongan tersebut relevan, lalu hubungkan langsung gap dengan tujuan dan kontribusi penelitian ini."
    if "desain" in lab or "pendekatan" in lab:
        return f"Nyatakan secara eksplisit desain/pendekatan yang benar-benar digunakan, alasan pemilihannya, dan hubungan desain tersebut dengan tujuan penelitian. Untuk {rt}, istilah desain harus konsisten dengan sumber data, teknik analisis, dan bentuk hasil yang dilaporkan."
    if "subjek" in lab or "sumber data" in lab or "objek" in lab:
        return "Jelaskan unit analisis atau sumber data secara operasional: siapa/apa yang diteliti, jumlah atau cakupan data, lokasi/periode bila relevan, serta kriteria pemilihan. Informasi harus memungkinkan pembaca menilai kecukupan sumber data terhadap tujuan penelitian."
    if "pengumpulan" in lab:
        return "Uraikan teknik dan prosedur pengumpulan data secara berurutan, termasuk instrumen/sumber data, pelaksanaan, periode, dan langkah pengendalian kualitas bila relevan. Gunakan hanya prosedur yang benar-benar dilakukan dan dapat dipertanggungjawabkan."
    if "analisis" in lab:
        return "Jelaskan teknik analisis, tahapan pengolahan data, kriteria atau dasar interpretasi, dan alasan pemilihannya. Pastikan teknik analisis dapat menjawab tujuan/research question serta menghasilkan temuan yang kemudian dilaporkan pada bagian hasil."
    if "validitas" in lab or "trustworthiness" in lab:
        return "Laporkan strategi validitas, reliabilitas, atau trustworthiness yang benar-benar diterapkan, termasuk prosedur, kriteria, dan hasil penerapannya bila tersedia. Sesuaikan dengan desain penelitian; jangan menambahkan uji atau prosedur yang tidak dilakukan."
    if "etika" in lab:
        return "Jelaskan perlindungan partisipan dan data sesuai karakter penelitian, termasuk persetujuan partisipan, kerahasiaan, persetujuan etik bila diwajibkan, dan penanganan risiko. Uraian harus mencerminkan prosedur yang benar-benar diterapkan."
    if "hasil" in lab:
        return "Sajikan temuan utama secara terstruktur dan hubungkan setiap temuan dengan data atau keluaran analisis yang menjadi dasarnya. Hindari memasukkan interpretasi atau klaim yang belum memiliki evidensi pada bagian hasil."
    if "perbandingan" in lab:
        return "Bandingkan setiap temuan utama dengan penelitian terdahulu yang paling relevan. Jelaskan persamaan atau perbedaan dan berikan alasan ilmiah berdasarkan teori, karakteristik data, metode, atau temuan empiris; jangan berhenti pada pernyataan 'sejalan' atau 'berbeda'."
    if "implikasi" in lab or "kontribusi" in lab:
        return "Nyatakan kontribusi atau implikasi yang benar-benar dapat diturunkan dari temuan. Bedakan kontribusi teoritis, metodologis, praktis, atau kebijakan dan hindari klaim yang lebih luas daripada bukti penelitian."
    if sec=="References":
        return "Audit setiap referensi berdasarkan relevansi terhadap argumen, fungsi sumber primer/sekunder, kemutakhiran, keterlacakan sitasi, dan kesesuaian gaya bibliografis. Pertahankan sumber klasik yang memang fundamental, tetapi jangan menggantinya semata-mata karena usia publikasi."
    if sec=="Judul":
        return "Pastikan judul merepresentasikan objek/fokus, ruang lingkup, dan karakter penelitian secara akurat. Judul tidak boleh memuat klaim, variabel, lokasi, atau desain yang tidak benar-benar didukung oleh naskah."
    if sec=="Abstrak":
        return "Sinkronkan abstrak dengan artikel lengkap: masalah/tujuan, metode, hasil utama, dan kontribusi harus berasal dari bagian yang sesuai. Hindari memasukkan temuan, angka, atau klaim yang tidak muncul dalam naskah utama."
    if sec=="Discussion":
        return "Bangun pembahasan dari temuan → interpretasi → perbandingan dengan literatur → penjelasan persamaan/perbedaan → implikasi/kontribusi. Setiap interpretasi harus memiliki dasar data atau literatur yang memadai."
    if sec=="Conclusion":
        return "Jawab tujuan atau research question secara langsung berdasarkan temuan utama. Kesimpulan harus menunjukkan implikasi yang proporsional, tidak mengulang seluruh pembahasan, dan tidak memperkenalkan data atau klaim baru."
    return "Perjelas unsur yang menjadi temuan pada bagian ini dengan bukti tekstual yang dapat ditelusuri, kemudian periksa kembali kesesuaiannya dengan tujuan, metode, hasil, dan kesimpulan."


def professional_alternatives(finding, section, text=""):
    """Three genuinely different revision paths, grounded in the finding."""
    problem=(finding.get("issue","") if isinstance(finding,dict) else str(finding)).lower()
    label=section
    keys=["desain/pendekatan","subjek/objek/sumber data","pengumpulan data","analisis data","validitas/reliabilitas atau trustworthiness","etika penelitian","perbandingan literatur","implikasi/kontribusi","gap","state of the art","tujuan","masalah","hasil","judul","abstrak"]
    for k in keys:
        if k in problem: label=k; break
    mapping={
        "tujuan":[
            "Alternatif A — Tambahkan satu pernyataan tujuan yang eksplisit pada akhir Pendahuluan dan pastikan setiap tujuan memperoleh jawaban pada Hasil serta Conclusion.",
            "Alternatif B — Jika tujuan sudah tersebar, konsolidasikan menjadi tujuan utama dan tujuan khusus tanpa menambah tujuan baru yang tidak dikerjakan.",
            "Alternatif C — Uji ulang tujuan terhadap judul dan metode; pertahankan hanya tujuan yang benar-benar dapat dijawab oleh desain dan data yang tersedia."],
        "masalah":[
            "Alternatif A — Tambahkan bukti empiris atau data yang menunjukkan fenomena yang menjadi masalah penelitian.",
            "Alternatif B — Susun ulang paragraf dari fenomena → konsekuensi ilmiah/praktis → kebutuhan penelitian.",
            "Alternatif C — Jika masalah sudah tersirat, pertegas kalimat masalah tanpa menambahkan klaim faktual baru."],
        "state of the art":[
            "Alternatif A — Sintesis beberapa studi primer berdasarkan tema/temuan sehingga terlihat posisi pengetahuan yang sudah tersedia.",
            "Alternatif B — Bandingkan pendekatan, data, atau temuan studi terdahulu dan tunjukkan keterbatasannya.",
            "Alternatif C — Buat paragraf penutup state of the art yang secara eksplisit mengantar pembaca menuju gap penelitian."],
        "gap":[
            "Alternatif A — Formulasikan gap berbasis perbedaan atau keterbatasan temuan penelitian terdahulu yang dapat diverifikasi.",
            "Alternatif B — Formulasikan gap berbasis keterbatasan pendekatan/metode yang digunakan studi sebelumnya.",
            "Alternatif C — Jika gap sudah ada tetapi tersebar, konsolidasikan menjadi satu argumen gap → tujuan → kontribusi."],
        "desain/pendekatan":[
            "Alternatif A — Nyatakan nama desain dan alasan pemilihannya pada awal Metodologi.",
            "Alternatif B — Jika desain sudah benar tetapi istilahnya kabur, susun ulang menjadi pendekatan → desain → tujuan penggunaannya.",
            "Alternatif C — Verifikasi konsistensi desain dengan sumber data, analisis, hasil, dan kesimpulan sebelum merevisi istilah."],
        "subjek/objek/sumber data":[
            "Alternatif A — Tambahkan identitas operasional sumber data, jumlah/cakupan, lokasi/periode, dan kriteria pemilihan.",
            "Alternatif B — Satukan informasi yang tersebar dalam subbagian khusus Sumber Data/Subjek/Unit Analisis.",
            "Alternatif C — Jelaskan alasan sumber data tersebut memadai untuk menjawab tujuan penelitian."],
        "pengumpulan data":[
            "Alternatif A — Tulis prosedur pengumpulan data secara kronologis dari persiapan sampai data siap dianalisis.",
            "Alternatif B — Sajikan teknik/instrumen, sumber data, periode, dan pelaksanaan dalam tabel prosedural bila jurnal mengizinkan.",
            "Alternatif C — Jika prosedur sudah dilakukan tetapi belum tertulis, rekonstruksi hanya dari catatan penelitian yang tersedia."],
        "analisis data":[
            "Alternatif A — Jelaskan teknik analisis dan urutan tahapannya sampai menghasilkan temuan.",
            "Alternatif B — Hubungkan setiap research question/tujuan dengan teknik analisis yang digunakan.",
            "Alternatif C — Audit kembali kesesuaian teknik analisis dengan jenis data dan hasil yang dilaporkan; hapus istilah analisis yang tidak digunakan."],
        "validitas/reliabilitas atau trustworthiness":[
            "Alternatif A — Laporkan prosedur validitas/reliabilitas atau trustworthiness beserta kriteria/hasilnya.",
            "Alternatif B — Untuk kualitatif, jelaskan strategi yang benar-benar diterapkan seperti triangulasi atau member checking.",
            "Alternatif C — Untuk kuantitatif, jelaskan instrumen, jenis pengujian, kriteria, dan hasil pengujiannya bila memang dilakukan."],
        "etika penelitian":[
            "Alternatif A — Tambahkan paragraf etika yang memuat persetujuan, kerahasiaan, dan perlindungan data/subjek sesuai penelitian.",
            "Alternatif B — Jika etika sudah tersebar, konsolidasikan seluruh prosedur yang benar-benar dilakukan.",
            "Alternatif C — Verifikasi apakah penelitian memerlukan persetujuan etik dan dokumentasikan statusnya secara faktual."],
        "perbandingan literatur":[
            "Alternatif A — Hubungkan setiap temuan utama dengan studi terdahulu yang paling relevan.",
            "Alternatif B — Jika hasil berbeda, jelaskan kemungkinan penyebab berdasarkan teori, metode, karakteristik data, atau kondisi empiris.",
            "Alternatif C — Ganti pernyataan deskriptif seperti 'sejalan'/'berbeda' dengan interpretasi yang menjelaskan makna ilmiahnya."],
        "implikasi/kontribusi":[
            "Alternatif A — Nyatakan kontribusi teoritis atau konseptual yang langsung berasal dari temuan.",
            "Alternatif B — Nyatakan implikasi praktis/kebijakan hanya jika didukung hasil dan ruang lingkup penelitian.",
            "Alternatif C — Jika klaim terlalu luas, persempit kontribusi agar proporsional dengan data dan desain penelitian."],
        "hasil":[
            "Alternatif A — Susun hasil mengikuti urutan tujuan/research question sehingga keterlacakan jawaban lebih mudah.",
            "Alternatif B — Hubungkan setiap temuan dengan tabel, gambar, kutipan, ukuran statistik, atau evidensi lain yang tersedia.",
            "Alternatif C — Pisahkan temuan empiris dari interpretasi yang seharusnya ditempatkan pada Discussion."],
        "judul":[
            "Alternatif A — Pertahankan fokus utama dan padatkan kata yang tidak memberi informasi ilmiah.",
            "Alternatif B — Revisi judul berdasarkan objek/fokus dan ruang lingkup yang benar-benar tercermin dalam data.",
            "Alternatif C — Uji judul terhadap tujuan, metode, hasil, dan kontribusi; hapus unsur yang tidak didukung naskah."],
        "abstrak":[
            "Alternatif A — Susun ulang menjadi masalah/tujuan → metode → hasil utama → kontribusi/implikasi.",
            "Alternatif B — Gunakan kalimat tujuan dan hasil yang lebih spesifik berdasarkan isi naskah, bukan generalisasi.",
            "Alternatif C — Audit abstrak baris demi baris terhadap Pendahuluan, Metodologi, Hasil, dan Conclusion."],
    }
    if label in mapping: return mapping[label]
    if section=="References":
        return [
            "Alternatif A — Audit keterhubungan sitasi dalam teks dengan daftar pustaka dan sebaliknya.",
            "Alternatif B — Evaluasi relevansi serta fungsi sumber primer, review/sekunder, buku, dan dokumen non-jurnal.",
            "Alternatif C — Perbarui sumber mutakhir yang relevan sambil mempertahankan sumber klasik yang memang menjadi landasan fundamental."
        ]
    return [
        f"Alternatif A — Perbaiki langsung unsur yang menjadi temuan pada bagian {section} berdasarkan bukti yang telah diverifikasi.",
        f"Alternatif B — Jika unsur sebenarnya sudah ada, pertegas lokasinya dan susun ulang uraian agar hubungan logisnya mudah ditelusuri.",
        f"Alternatif C — Setelah revisi, lakukan pemeriksaan silang {section} dengan tujuan, metode, hasil, dan kesimpulan agar tidak muncul ketidakkonsistenan baru."
    ]


def analyze_section_demo(section,text):
    if section=="References":
        return analyze_reference_demo(text, split_sections(st.session_state.article_text).get("Judul","") if st.session_state.get("article_text") else "")
    words=len(text.split()); findings=[]; recommendations=[]
    title=split_sections(st.session_state.article_text).get("Judul","") if st.session_state.get("article_text") else ""
    rtype=detect_research_type(st.session_state.get("article_text","") or text)

    def add(label,terms,severity,problem,action):
        ev=find_evidence(text,terms)
        if ev and any(t.lower() in ev.lower() for t in terms):
            evidence=evidence_for_present(text,terms,label,section)
        else:
            evidence=evidence_for_missing(text,terms,label,section)
        f=issue(severity,section,problem,evidence,action)
        f["alternatives"]=professional_alternatives(f,section,text)
        f["reviewer_recommendation"]=professional_recommendation(section,label,text,rtype)
        findings.append(f)

    if section=="Judul":
        if words<5: add("judul",[" "] ,"MINOR","Judul sangat singkat sehingga fokus penelitian belum dapat dibaca secara memadai.","Pastikan judul menyebut objek/fokus utama dan karakter penelitian yang memang relevan.")
        if len(text)>220: add("judul",[text[:30]],"MINOR","Judul relatif panjang dan berpotensi mengaburkan fokus utama.","Padatkan unsur yang tidak esensial tanpa menghilangkan objek, fokus, atau desain yang relevan.")
        if title:
            findings.append(issue("EDITORIAL",section,"Judul terdeteksi sebagai satu kesatuan dan akan ditampilkan utuh pada kartu judul.",f'Judul terbaca: “{title}”',"Periksa keselarasan judul dengan tujuan, metode, hasil, dan kontribusi."))
        recommendations.append("Nilai judul berdasarkan empat unsur: objek/fokus, ruang lingkup, pendekatan atau variabel bila memang diperlukan, dan kontribusi yang tercermin secara akurat. Jangan memasukkan istilah metodologis yang tidak benar-benar digunakan.")
    elif section=="Abstrak":
        for label,terms in [("tujuan",["tujuan","bertujuan","aim","objective"]),("metode",["metode","method","pendekatan","approach"]),("hasil",["hasil","temuan","result","finding"]),("kontribusi",["kontribusi","implikasi","contribution","implication"])]:
            if not contains_any(text,terms): add(label,terms,"MINOR",f"Komponen {label} belum terdeteksi secara tekstual.",f"Nyatakan {label} secara eksplisit dan selaraskan dengan isi artikel.")
        recommendations.append("Susun abstrak sebagai ringkasan padat yang memuat masalah/tujuan, metode, hasil utama berbasis data, dan kontribusi/implikasi. Hindari klaim yang tidak muncul pada bagian hasil.")
    elif section=="Pendahuluan":
        for label,terms in [("masalah",["masalah","problem","fenomena","isu"]),("state of the art",["penelitian terdahulu","studi terdahulu","literatur","state of the art"]),("gap",["research gap","kesenjangan","belum","masih terbatas"]),("tujuan",["tujuan penelitian","bertujuan","aim","objective"])]:
            if not contains_any(text,terms): add(label,terms,"MAJOR" if label in ["masalah","gap","tujuan"] else "MINOR",f"Elemen {label} belum terdeteksi secara tekstual.",f"Tunjukkan {label} secara eksplisit dan hubungkan dengan paragraf sebelum dan sesudahnya.")
        recommendations.append("Bangun argumentasi secara bertahap dari fenomena/masalah → state of the art → keterbatasan penelitian terdahulu → gap yang spesifik → tujuan/research question → kontribusi. Setiap klaim tentang penelitian terdahulu harus ditopang sitasi yang relevan.")
    elif section=="Metodologi":
        checks=[("desain/pendekatan",["pendekatan","desain","design","method","metode"],"MAJOR"),("subjek/objek/sumber data",["subjek","objek","informan","responden","sampel","populasi","sumber data"],"MAJOR"),("pengumpulan data",["wawancara","observasi","dokumentasi","kuesioner","pengumpulan data","data collection"],"MAJOR"),("analisis data",["analisis data","data analysis","coding","analisis tematik","regresi","anova","uji t","uji f"],"MAJOR"),("validitas/reliabilitas atau trustworthiness",["validitas","reliabilitas","reliability","validity","trustworthiness","credibility","triangulasi"],"MINOR"),("etika penelitian",["etika penelitian","ethical","informed consent","persetujuan etik","komite etik"],"MINOR")]
        for label,terms,sev in checks:
            if not contains_any(text,terms): add(label,terms,sev,f"Komponen {label} belum terdeteksi secara tekstual.",f"Lengkapi {label} secara operasional sesuai desain penelitian yang benar-benar dilakukan.")
        recommendations.append(recommendation_for_methodology(text,rtype))
    elif section=="Hasil":
        if not contains_any(text,["tabel","table","gambar","figure","temuan","menunjukkan","hasil penelitian"]): add("hasil",["tabel","table","gambar","figure","temuan","menunjukkan","hasil penelitian"],"MAJOR","Evidensi hasil belum terdeteksi secara tekstual dengan kuat.","Tampilkan temuan utama beserta data/evidensi yang memungkinkan pembaca menelusuri dasar kesimpulan.")
        recommendations.append("Pisahkan penyajian temuan dari interpretasi yang terlalu dini. Setiap temuan utama harus dapat ditelusuri ke data atau hasil analisis yang dilaporkan, dengan tabel/gambar/kutipan/ukuran statistik sesuai desain penelitian.")
    elif section=="Discussion":
        if not contains_any(text,["penelitian terdahulu","studi terdahulu","sejalan","berbeda","dibandingkan","literatur","previous studies"]): add("perbandingan literatur",["penelitian terdahulu","studi terdahulu","sejalan","berbeda","dibandingkan","literatur","previous studies"],"MAJOR","Perbandingan hasil dengan penelitian terdahulu belum terdeteksi secara tekstual.","Bandingkan temuan dengan studi relevan dan jelaskan alasan persamaan/perbedaannya.")
        if not contains_any(text,["implikasi","kontribusi","contribution","implication"]): add("implikasi/kontribusi",["implikasi","kontribusi","contribution","implication"],"MINOR","Implikasi atau kontribusi belum tampak kuat secara tekstual.","Jelaskan kontribusi teoritis, metodologis, praktis, atau kebijakan sesuai bukti penelitian.")
        recommendations.append("Discussion harus menjawab pertanyaan ‘apa arti temuan ini?’. Hubungkan temuan dengan teori dan studi terdahulu, jelaskan persamaan/perbedaan, berikan alasan ilmiah yang didukung data/literatur, lalu turunkan implikasi dan kontribusi tanpa melampaui bukti.")
    elif section=="Conclusion":
        if words<80: add("kesimpulan",["kesimpulan","simpulan","berdasarkan","hasil"],"MINOR","Kesimpulan relatif singkat; kecukupannya perlu dinilai terhadap tujuan dan research question.","Pastikan setiap tujuan/research question memperoleh jawaban yang bersumber dari temuan penelitian.")
        recommendations.append("Kesimpulan harus menjawab tujuan/research question berdasarkan hasil utama, menyatakan kontribusi yang benar-benar didukung data, serta menghindari data atau klaim baru yang tidak dibahas sebelumnya.")
    if not findings:
        f=issue("EDITORIAL",section,"Tidak ditemukan masalah struktural dominan pada pemeriksaan awal.",f"Status bukti: TERKONFIRMASI. Indikator utama pada {section} memiliki jejak tekstual yang dapat ditelusuri; pemeriksaan awal belum menemukan kekosongan struktural dominan.","Lanjutkan pemeriksaan substantif dan konsistensi antarbagian.")
        f["alternatives"]=professional_alternatives(f,section,text)
        f["reviewer_recommendation"]=professional_recommendation(section,"",text,rtype)
        findings.append(f)
    # Recommendations are now derived from the actual findings, preventing one generic sentence from appearing for every article.
    recommendations=[]
    for f in findings:
        if isinstance(f,dict):
            rec=f.get("reviewer_recommendation") or professional_recommendation(section,f.get("issue",""),text,rtype)
            if rec and rec not in recommendations: recommendations.append(rec)
    return {"section":section,"words":words,"text":text,"findings":findings,"recommendations":recommendations}


def parse_reference_entries(text):
    """Parse bibliography into coarse entries while preserving original text.
    This is a screening aid, not a claim about the true source type.
    """
    if not text.strip(): return []
    raw=[x.strip() for x in text.splitlines() if x.strip()]
    entries=[]; current=""
    for line in raw:
        starts=bool(re.match(r"^(?:\[?\d+\]?|\d+\.|[-•])\s+",line))
        year=bool(re.search(r"\b(?:19|20)\d{2}[a-z]?\b",line))
        if starts and current:
            entries.append(current.strip()); current=line
        elif starts:
            current=line
        elif current:
            current += " " + line
        elif year:
            entries.append(line)
        else:
            current=line
    if current: entries.append(current.strip())
    if len(entries)==1 and len(raw)>1:
        # Fall back to paragraph-like blocks when the source has no numbering.
        blocks=[b.strip() for b in re.split(r"\n\s*\n",text) if b.strip()]
        if len(blocks)>1: entries=blocks
    return entries


def classify_reference(entry):
    low=entry.lower()
    if any(x in low for x in ["doi.org/","journal","jurnal","volume","vol.","no.","issn"]):
        if any(x in low for x in ["systematic review","literature review","scoping review","meta-analysis","meta analysis","review article"]):
            return "Jurnal Sekunder/Review"
        return "Jurnal/Artikel Kandidat Primer"
    if any(x in low for x in ["publisher","penerbit","press","edition","edisi","isbn"]): return "Buku"
    if any(x in low for x in ["thesis","dissertation","skripsi","tesis","disertasi"]): return "Skripsi/Tesis/Disertasi"
    if any(x in low for x in ["report","laporan","policy brief","working paper"]): return "Laporan/Dokumen"
    if re.search(r"https?://",entry): return "Web/Daring"
    return "Lainnya"


def reference_audit(text,title=""):
    entries=parse_reference_entries(text)
    rows=[]
    title_terms=set(re.findall(r"[a-zA-ZÀ-ÿ]{5,}", (title or "").lower()))
    stop={"penelitian","research","study","studi","analisis","analysis","dengan","untuk","dalam","the","and","yang","this","based"}
    title_terms={x for x in title_terms if x not in stop}
    for i,e in enumerate(entries,1):
        years=re.findall(r"\b(?:19|20)\d{2}\b",e)
        year=int(years[-1]) if years else None
        words=set(re.findall(r"[a-zA-ZÀ-ÿ]{5,}",(e.lower())))
        overlap=len(title_terms & words) if title_terms else 0
        rows.append({"No":i,"Jenis":classify_reference(e),"Tahun":year or "-","Kecocokan kata judul":overlap,"Referensi":e})
    return rows


def count_reference_entries(text):
    """Estimate bibliography entries from common numbering/author/year patterns."""
    if not text.strip(): return 0
    lines=[x.strip() for x in text.splitlines() if x.strip()]
    count=0
    for line in lines:
        if re.match(r"^(?:\[?\d+\]?|[-•])\s+",line):
            count+=1
        elif re.search(r"\b(?:19|20)\d{2}[a-z]?\b",line) and len(line.split())>=4:
            count+=1
    return count or len(lines)


def analyze_reference_demo(text,title=""):
    words=len(text.split()); years=[int(y) for y in re.findall(r"\b(19\d{2}|20\d{2})\b",text)]
    entries=count_reference_entries(text); rows=reference_audit(text,title); findings=[]; strengths=[]; recommendations=[]
    def add_ref_finding(severity, problem, terms, action, label="referensi"):
        evidence=evidence_for_present(text,terms,label,"References") if find_evidence(text,terms) else evidence_for_missing(text,terms,label,"References")
        f=issue(severity,"References",problem,evidence,action)
        f["alternatives"]=professional_alternatives(f,"References",text)
        f["reviewer_recommendation"]=professional_recommendation("References",problem,text,detect_research_type(st.session_state.get("article_text","") or text))
        findings.append(f)
    if not text.strip():
        add_ref_finding("MAJOR","Daftar pustaka belum berhasil diekstraksi dari naskah.",["daftar pustaka","referensi","references"],"Periksa heading atau masukkan bagian referensi secara manual.")
    else:
        journal=sum(1 for r in rows if r["Jenis"] in ["Jurnal/Artikel Kandidat Primer","Jurnal Sekunder/Review"])
        nonjournal=len(rows)-journal
        if entries<5:
            add_ref_finding("MAJOR",f"Entri referensi yang terdeteksi relatif sedikit, sekitar {entries} entri.",["19","20"],"Verifikasi kelengkapan daftar pustaka terhadap seluruh sitasi dalam teks.","referensi")
        if not years:
            add_ref_finding("MAJOR","Tahun publikasi belum terdeteksi secara memadai.",["19","20"],"Periksa kembali format setiap entri dan lengkapi tahun publikasi berdasarkan sumber aslinya.","referensi")
        else:
            recent=sum(1 for y in years if y>=2021); old=sum(1 for y in years if y<2016)
            if recent==0:
                add_ref_finding("MINOR","Belum terdeteksi sumber tahun 2021–sekarang.",["2021","2022","2023","2024","2025","2026"],"Tambahkan sumber primer mutakhir yang benar-benar relevan; jangan mengganti sumber klasik yang masih fundamental.","referensi")
            if old>recent*2 and len(years)>=5:
                add_ref_finding("MINOR","Proporsi sumber lama tampak dominan dibanding sumber 2021–sekarang.",["2015","2014","2013","2012","2011"],"Evaluasi fungsi sumber lama dan lengkapi literatur primer mutakhir yang relevan.","referensi")
        if title and rows:
            lowmatch=sum(1 for r in rows if r["Kecocokan kata judul"]>0)
            if lowmatch==0:
                add_ref_finding("MAJOR","Tidak ditemukan kecocokan kata kunci judul dengan entri referensi secara tekstual; hasil ini adalah sinyal screening dan bukan putusan relevansi substantif.",list(re.findall(r"[A-Za-zÀ-ÿ]{5,}",title.lower()))[:8],"Periksa relevansi substantif referensi terhadap fokus, teori, metode, dan temuan artikel secara manual/AI.","referensi")
        strengths.append(f"Bagian referensi berhasil dibaca secara tekstual; sekitar {entries} entri terdeteksi ({journal} jurnal/artikel dan {nonjournal} non-jurnal berdasarkan indikator format).")
        recommendations.extend([
            "Audit keterhubungan sitasi dalam teks dengan daftar pustaka dan sebaliknya.",
            "Prioritaskan sumber primer yang paling relevan dengan objek, fokus, teori, metode, dan temuan penelitian.",
            "Jangan memperlakukan jurnal review/meta-analisis otomatis sebagai bukti primer.",
            "Pertahankan sumber klasik yang fundamental dan lengkapi dengan penelitian primer mutakhir yang relevan.",
            "Pisahkan fungsi sumber jurnal, review/sekunder, buku, laporan, dan sumber daring secara jelas.",
            "Seragamkan gaya bibliografis sesuai pedoman jurnal target."
        ])
    if not findings:
        f=issue("EDITORIAL","References","Tidak ditemukan masalah struktural dominan pada audit awal referensi.",f"Status bukti: TERKONFIRMASI. Sekitar {entries} entri referensi dapat dibaca dan diklasifikasikan secara heuristik. Pemeriksaan jenis sumber dan relevansi tetap memerlukan verifikasi substantif.","Lanjutkan audit sitasi, relevansi sumber, kemutakhiran, dan gaya bibliografis.")
        f["alternatives"]=professional_alternatives(f,"References",text)
        f["reviewer_recommendation"]=professional_recommendation("References","",text,detect_research_type(st.session_state.get("article_text","") or text))
        findings.append(f)
    return {"section":"References","words":words,"text":text,"findings":findings,"strengths":strengths,"recommendations":recommendations,"entries":entries,"rows":rows}


def alternative_solutions(finding, section):
    """Return concrete alternative revision paths tied to the actual finding."""
    if not isinstance(finding, dict):
        text=str(finding)
        return [
            f"Jika unsur yang dimaksud memang ada dalam penelitian, jelaskan secara eksplisit pada bagian {section} dengan istilah yang konsisten dengan naskah.",
            f"Jika unsur tersebut belum dilakukan, revisi hanya dengan prosedur yang benar-benar dapat dipertanggungjawabkan; jangan menambahkan data atau prosedur baru secara fiktif.",
        ]
    problem=(finding.get("issue") or "").lower()
    action=(finding.get("action") or "").strip()
    label=section
    for key in ["desain/pendekatan","subjek/objek/sumber data","teknik pengumpulan data","teknik analisis data","validitas/reliabilitas atau trustworthiness","etika penelitian","perbandingan literatur","implikasi/kontribusi","gap","state of the art","tujuan"]:
        if key in problem:
            label=key
            break
    if section=="Metodologi":
        if "desain" in label or "pendekatan" in label:
            return [
                "Alternatif 1 — Nyatakan nama desain penelitian secara eksplisit, lalu jelaskan alasan desain tersebut paling sesuai dengan tujuan dan pertanyaan penelitian.",
                "Alternatif 2 — Jika desain sebenarnya sudah digunakan tetapi istilahnya belum jelas, rapikan paragraf metode dengan urutan pendekatan → desain → alasan pemilihan → implikasi terhadap analisis.",
                "Alternatif 3 — Cocokkan kembali desain dengan data dan hasil; hapus istilah desain yang tidak benar-benar digunakan."
            ]
        if "subjek" in label or "sumber data" in label:
            return [
                "Alternatif 1 — Tambahkan siapa/apa yang menjadi sumber data, jumlahnya, karakteristiknya, lokasi, dan kriteria pemilihannya.",
                "Alternatif 2 — Jika informasi sudah tersebar di beberapa paragraf, satukan dalam satu subbagian sumber data/subjek agar mudah ditelusuri.",
                "Alternatif 3 — Hubungkan sumber data dengan tujuan penelitian dan jelaskan alasan pemilihannya."
            ]
        if "pengumpulan" in label:
            return [
                "Alternatif 1 — Uraikan instrumen dan langkah pengumpulan data secara berurutan, termasuk waktu dan pelaksanaannya.",
                "Alternatif 2 — Jika prosedur sudah dilakukan tetapi belum terdokumentasi, tambahkan uraian berdasarkan catatan penelitian yang benar-benar tersedia.",
                "Alternatif 3 — Jelaskan hubungan setiap teknik pengumpulan data dengan jenis informasi yang ingin diperoleh."
            ]
        if "analisis" in label:
            return [
                "Alternatif 1 — Jelaskan teknik analisis, tahapan, kriteria keputusan, dan alasan pemilihannya sesuai jenis data.",
                "Alternatif 2 — Jika analisis sudah dilakukan, tuliskan kembali alurnya dari data mentah → pengolahan → analisis → keluaran yang menjadi dasar hasil.",
                "Alternatif 3 — Cocokkan teknik analisis dengan pertanyaan penelitian dan pastikan hasil yang disajikan benar-benar berasal dari teknik tersebut."
            ]
        if "validitas" in label or "trustworthiness" in label:
            return [
                "Alternatif 1 — Laporkan teknik validitas/reliabilitas atau strategi trustworthiness yang benar-benar digunakan beserta hasil atau prosedurnya.",
                "Alternatif 2 — Untuk penelitian kualitatif, jelaskan strategi seperti triangulasi, member checking, credibility, dependability, atau confirmability hanya jika benar-benar dilakukan.",
                "Alternatif 3 — Jika penelitian kuantitatif, jelaskan jenis uji validitas/reliabilitas, instrumen, kriteria, dan hasil pengujiannya."
            ]
        if "etika" in label:
            return [
                "Alternatif 1 — Tambahkan informasi persetujuan partisipan, kerahasiaan identitas/data, dan persetujuan etik bila diwajibkan.",
                "Alternatif 2 — Jika aspek etika sudah dilakukan tetapi tersebar, rangkum prosedurnya dalam satu paragraf etika penelitian.",
                "Alternatif 3 — Pastikan uraian etika sesuai dengan karakter subjek, data, lokasi, dan risiko penelitian."
            ]
    if section=="Pendahuluan" and ("gap" in label or "state of the art" in label):
        return [
            "Alternatif 1 — Tambahkan sintesis beberapa penelitian primer yang menunjukkan apa yang sudah diketahui dan apa yang belum terjawab.",
            "Alternatif 2 — Buat perbandingan eksplisit antarkajian terdahulu lalu tunjukkan keterbatasan yang langsung membuka ruang penelitian ini.",
            "Alternatif 3 — Tutup alur dengan kalimat gap → tujuan → kontribusi sehingga alasan penelitian dapat ditelusuri."
        ]
    if section=="Discussion" and "perbandingan" in label:
        return [
            "Alternatif 1 — Bandingkan setiap temuan utama dengan minimal studi terdahulu yang paling relevan.",
            "Alternatif 2 — Jika hasil berbeda, jelaskan kemungkinan penyebab berdasarkan teori, karakteristik data, metode, atau konteks empiris.",
            "Alternatif 3 — Hindari sekadar menyatakan 'sejalan' atau 'berbeda'; jelaskan makna ilmiah dari persamaan/perbedaan tersebut."
        ]
    if section=="References":
        return [
            "Alternatif 1 — Pertahankan sumber klasik yang fundamental dan tambahkan sumber primer mutakhir yang langsung relevan dengan fokus artikel.",
            "Alternatif 2 — Audit satu per satu: setiap sitasi dalam teks harus memiliki entri daftar pustaka dan setiap entri yang digunakan harus relevan dengan argumen.",
            "Alternatif 3 — Pisahkan sumber primer, review/sekunder, buku, dan dokumen non-jurnal agar fungsi setiap sumber jelas."
        ]
    if section=="Judul":
        return [
            "Alternatif 1 — Padatkan judul dengan mempertahankan objek/fokus utama dan ruang lingkup penelitian.",
            "Alternatif 2 — Hapus kata yang tidak menambah informasi ilmiah, tetapi jangan menghilangkan unsur yang membedakan penelitian.",
            "Alternatif 3 — Uji kembali keselarasan judul dengan tujuan, metode, hasil, dan kontribusi."
        ]
    if action:
        return [
            f"Alternatif 1 — {action}",
            f"Alternatif 2 — Jika unsur tersebut sebenarnya sudah dilakukan, pindahkan atau pertegas uraian yang relevan pada bagian {section} tanpa menambah fakta baru.",
            f"Alternatif 3 — Verifikasi kembali revisi terhadap tujuan, data, hasil, dan kesimpulan agar perbaikan tidak menimbulkan ketidakkonsistenan baru."
        ]
    return [
        f"Alternatif 1 — Perbaiki langsung bagian {section} berdasarkan temuan dan bukti naskah.",
        f"Alternatif 2 — Jika unsur sebenarnya sudah ada, pertegas dan susun ulang uraian agar dapat ditelusuri reviewer.",
        "Alternatif 3 — Lakukan pemeriksaan silang dengan bagian lain sebelum menetapkan revisi akhir."
    ]


def finding_text(finding):
    if isinstance(finding, dict):
        return finding.get("issue", "")
    return str(finding)


def render_alternatives(finding, section):
    st.markdown("**Alternatif solusi**")
    for i, sol in enumerate(alternative_solutions(finding, section), 1):
        st.markdown(f'<div class="solution-box"><b>{i}.</b> {html.escape(sol)}</div>', unsafe_allow_html=True)


def render_section_analysis(result):
    st.markdown("### Hasil Analisis Bagian")
    st.write(f"**Bagian:** {result['section']} · **{result['words']:,} kata**")
    st.markdown("**Temuan reviewer**")
    findings=result.get("findings",[]) or []
    for i,x in enumerate(findings,1):
        if isinstance(x,dict):
            sev=x.get("severity","MINOR")
            section_name=x.get("section",result["section"])
            problem=x.get("issue","")
            ev=x.get("evidence","") or "Dasar pemeriksaan belum tersedia."
            rec=x.get("reviewer_recommendation") or x.get("action","")
            st.markdown(
                f'<div class="finding-box">'
                f'<div style="font-size:16px;font-weight:700;margin-bottom:7px">{i}. {html.escape(sev)} · {html.escape(section_name)}</div>'
                f'<div><b>Temuan:</b> {html.escape(problem)}</div>'
                f'<div class="evidence-box"><b>📌 Dasar/Bukti pemeriksaan</b><br>{html.escape(ev)}</div>'
                f'<div class="recommendation-box"><b>Rekomendasi reviewer:</b><br>{html.escape(rec)}</div>'
                f'</div>',unsafe_allow_html=True
            )
            sols=x.get("alternatives") or professional_alternatives(x,result["section"],result.get("text",""))
            st.markdown("**Alternatif solusi**")
            for j,sol in enumerate(sols,1):
                st.markdown(f'<div class="solution-box"><b>{chr(64+j)}.</b> {html.escape(sol)}</div>',unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="finding-box"><b>{i}.</b> {html.escape(str(x))}</div>',unsafe_allow_html=True)
    # Recommendations shown here are a concise index; detailed recommendations stay attached to each finding.
    if result.get("recommendations"):
        st.markdown("### Ringkasan rekomendasi revisi")
        for i,x in enumerate(result["recommendations"],1):
            st.markdown(f'<div class="recommendation-box"><b>{i}.</b> {html.escape(str(x))}</div>',unsafe_allow_html=True)

def render_reference_table(rows):
    """Render a wrapped fixed-layout reference table without horizontal scrolling."""
    if not rows:
        st.info("Belum ada entri referensi yang dapat ditampilkan.")
        return
    head='<div class="review-table-wrap"><table class="review-table"><thead><tr><th class="c-no">No</th><th class="c-type">Jenis</th><th class="c-year">Tahun</th><th class="c-match">Kecocokan</th><th class="c-ref">Referensi</th></tr></thead><tbody>'
    body=[]
    for r in rows:
        body.append('<tr>'
                    f'<td class="c-no">{html.escape(str(r.get("No","")))}</td>'
                    f'<td class="c-type">{html.escape(str(r.get("Jenis","")))}</td>'
                    f'<td class="c-year">{html.escape(str(r.get("Tahun","-")))}</td>'
                    f'<td class="c-match">{html.escape(str(r.get("Kecocokan kata judul",0)))}</td>'
                    f'<td class="c-ref">{html.escape(str(r.get("Referensi","")))}</td>'
                    '</tr>')
    st.markdown(head+''.join(body)+'</tbody></table></div>',unsafe_allow_html=True)


def _ai_call(prompt, system, key, endpoint, model):
    resp=requests.post(
        endpoint.rstrip("/")+"/chat/completions",
        headers={"Authorization":f"Bearer {key}","Content-Type":"application/json"},
        json={"model":model,"messages":[{"role":"system","content":system},{"role":"user","content":prompt}],"temperature":0.1},
        timeout=240
    )
    resp.raise_for_status()
    content=resp.json()["choices"][0]["message"]["content"]
    content=re.sub(r"^```(?:json)?\s*|\s*```$","",content.strip())
    return json.loads(content)


def ai_review(text,meta,key,endpoint,model):
    """Review the complete article. For long articles, review chunks first and synthesize them."""
    system="""Anda adalah reviewer jurnal ilmiah senior, editor, metodolog, dan akademisi. Review artikel secara objektif, berbasis bukti dari teks, tanpa mengarang. Bedakan temuan, inferensi terbatas, dan rekomendasi. Jangan menyatakan novelty/gap terbukti bila literatur primer belum dibandingkan. Jika informasi tidak tersedia, tulis 'tidak ditemukan dalam artikel' dan jangan mengisinya dengan asumsi."""
    base_rules="""
PRINSIP: baca seluruh materi yang diberikan; jangan mengabaikan bagian akhir; jangan menyimpulkan berdasarkan potongan awal saja.
Audit Judul, Abstrak, Pendahuluan, state of the art, research gap, novelty, metodologi, hasil, discussion, conclusion, references, writing, ethics.
Uji konsistensi: Judul → Tujuan → Research Question → Metode → Data → Hasil → Discussion → Conclusion → Abstract.
Setiap isu harus memiliki severity CRITICAL/MAJOR/MINOR/EDITORIAL, bagian, masalah, bukti, dan tindakan revisi. Field evidence WAJIB menunjuk teks artikel: gunakan kutipan verbatim pendek (maksimal sekitar 300 karakter) yang benar-benar terdapat pada materi yang diberikan, atau jelaskan secara eksplisit bahwa indikator tidak ditemukan pada bagian yang diperiksa. Jangan membuat kutipan atau nomor halaman yang tidak tersedia.
Research gap dan novelty tidak boleh dinyatakan terbukti tanpa dasar komparatif yang tersedia.
Jika artikel kurang lengkap, nyatakan keterbatasan secara eksplisit.
Untuk referensi: prioritaskan sumber primer yang relevan; bedakan kandidat sumber primer, sumber sekunder/review, dan sumber non-jurnal; jangan menyebut jurnal sebagai primer hanya karena formatnya jurnal; jangan memasukkan sumber yang tidak relevan dengan judul/fokus aktif.
"""

    # Keep a safe per-request budget while ensuring the whole article is covered.
    chunk_size=50000
    chunks=[text[i:i+chunk_size] for i in range(0,len(text),chunk_size)] or [""]
    chunk_reviews=[]
    if len(chunks)==1:
        prompt=f"""Lakukan FULL REVIEW terhadap artikel berikut.\n{base_rules}\nMETADATA: {json.dumps(meta,ensure_ascii=False)}\n\nKEMBALIKAN JSON VALID SAJA dengan struktur:\n{{"summary":"...","research_type":"...","gap":"...","novelty":"...","method":"...","consistency":"...","issues":[{{"severity":"MAJOR","section":"Metodologi","issue":"...","evidence":"...","action":"..."}}],"score":0,"recommendation":"MAJOR REVISION","roadmap":["..."],"metrics":{{"section_scores":{{"Title & Abstract":0,"Introduction & Gap":0,"Literature & Framework":0,"Methodology":0,"Results":0,"Discussion":0,"Conclusion":0,"References":0,"Writing":0}}}}}}\n\nARTIKEL LENGKAP (awal sampai akhir):\n{text}"""
        return _ai_call(prompt,system,key,endpoint,model)

    # Long-document path: every chunk is reviewed, then the findings are synthesized.
    for i,chunk in enumerate(chunks,1):
        prompt=f"""Analisis bagian {i} dari {len(chunks)} artikel. Ini bukan review final; ekstrak semua bukti yang relevan dari bagian ini dan jangan mengarang informasi yang tidak ada.\n{base_rules}\nKembalikan JSON valid: {{"chunk_summary":"...","research_type_signals":"...","evidence":[{{"section":"...","finding":"...","evidence":"...","severity":"CRITICAL|MAJOR|MINOR|EDITORIAL"}}],"gap_signals":"...","novelty_signals":"...","method_signals":"...","consistency_signals":"..."}}\n\nTEKS BAGIAN {i}:\n{chunk}"""
        chunk_reviews.append(_ai_call(prompt,system,key,endpoint,model))

    synthesis=f"""Susun FULL REVIEW berdasarkan hasil pembacaan seluruh {len(chunks)} bagian artikel berikut. Jangan menganggap bagian yang tidak disebut sebagai tidak ada; gunakan semua evidence dari seluruh chunk. {base_rules}

METADATA: {json.dumps(meta,ensure_ascii=False)}

KEMBALIKAN JSON VALID SAJA dengan struktur:
{{"summary":"...","research_type":"...","gap":"...","novelty":"...","method":"...","consistency":"...","issues":[{{"severity":"MAJOR","section":"Metodologi","issue":"...","evidence":"...","action":"..."}}],"score":0,"recommendation":"MAJOR REVISION","roadmap":["..."],"metrics":{{"section_scores":{{"Title & Abstract":0,"Introduction & Gap":0,"Literature & Framework":0,"Methodology":0,"Results":0,"Discussion":0,"Conclusion":0,"References":0,"Writing":0}}}}}}

HASIL PEMBACAAN SEMUA BAGIAN:
{json.dumps(chunk_reviews,ensure_ascii=False)}"""
    return _ai_call(synthesis,system,key,endpoint,model)


def validate_review_evidence(review, article_text):
    """Verify AI evidence against the uploaded manuscript and make the audit status explicit."""
    if not isinstance(review,dict): return review
    source=normalize_space(article_text).lower()
    sections=split_sections(article_text)
    for x in review.get("issues",[]) or []:
        if not isinstance(x,dict): continue
        section=x.get("section","") or ""
        ev=normalize_space(str(x.get("evidence", "")))
        if not ev: continue
        quotes=re.findall(r'[“"]([^”"]{20,520})[”"]', ev)
        candidates=quotes or [ev]
        verified_candidates=[c for c in candidates if len(normalize_space(c))>=20 and normalize_space(c).lower() in source]
        if verified_candidates:
            q=verified_candidates[0]
            label=x.get("issue","")
            # Reconstruct a concise, academically interpretable evidence note from the verified quote.
            x["evidence"]=(
                f"Status verifikasi: TERKONFIRMASI. Lokasi pemeriksaan: bagian {section or 'yang dilaporkan AI'}. "
                f"Cuplikan naskah verbatim: “{q}” "
                f"Kutipan tersebut menjadi dasar faktual bagi temuan reviewer; penilaian akademiknya tetap didasarkan pada hubungan kutipan dengan masalah yang diidentifikasi, bukan pada kemunculan kata kunci semata."
            )
        elif "tidak ditemukan" in ev.lower() or "belum terdeteksi" in ev.lower():
            sec_text=sections.get(section,"")
            # Recover likely indicators from the finding text for a more specific negative audit.
            low_issue=x.get("issue","").lower()
            label_map={
                "tujuan":["tujuan","bertujuan","aim","objective"],
                "masalah":["masalah","problem","fenomena","isu"],
                "gap":["research gap","kesenjangan","belum","masih terbatas"],
                "subjek/objek/sumber data":["subjek","objek","informan","responden","sampel","populasi","sumber data"],
                "analisis data":["analisis data","data analysis","coding","regresi","anova","uji t","uji f"],
                "validitas/reliabilitas atau trustworthiness":["validitas","reliabilitas","reliability","validity","trustworthiness","triangulasi"],
                "etika penelitian":["etika penelitian","ethical","informed consent","persetujuan etik","komite etik"],
                "perbandingan literatur":["penelitian terdahulu","studi terdahulu","sejalan","berbeda","dibandingkan","literatur"],
            }
            chosen_label=next((k for k in label_map if k in low_issue), section or "unsur penelitian")
            terms=label_map.get(chosen_label,[chosen_label])
            x["evidence"]=evidence_for_missing(sec_text,terms,chosen_label,section)
        else:
            x["evidence"]=(
                f"Status verifikasi: BELUM TERKONFIRMASI. Pernyataan reviewer diarahkan pada bagian {section or 'yang dilaporkan AI'}, tetapi bukti yang diberikan tidak dapat dicocokkan secara verbatim dengan naskah yang diunggah. "
                f"Temuan ini harus diperlakukan sebagai sinyal untuk pemeriksaan lanjutan, bukan sebagai fakta final. Reviewer perlu menemukan lokasi dan kutipan naskah sebelum menjadikannya dasar keputusan editorial."
            )
    return review


def build_staged_review():
    """Aggregate Review Bertahap results without inventing findings."""
    stages=st.session_state.get("stage_reviews",{}) or {}
    if not stages: return None
    issues=[]; roadmap=[]; section_rows=[]
    for sec in SECTIONS:
        r=stages.get(sec)
        if not r: continue
        fs=r.get("findings",[]) or []
        for f in fs:
            if isinstance(f,dict):
                f.setdefault("alternatives", professional_alternatives(f, sec, stages.get(sec,{}).get("text", "")))
                f.setdefault("reviewer_recommendation", professional_recommendation(sec, f.get("issue", ""), stages.get(sec,{}).get("text", ""), detect_research_type(st.session_state.get("article_text", ""))))
                issues.append(f)
        for rec in r.get("recommendations",[]) or []:
            roadmap.append(f"{sec}: {rec}")
        section_rows.append({"Bagian":sec,"Kata":r.get("words",0),"Temuan":sum(1 for f in fs if isinstance(f,dict) and f.get("severity")!="EDITORIAL")})
    severity_order={"CRITICAL":0,"MAJOR":1,"MINOR":2,"EDITORIAL":3}
    issues=sorted(issues,key=lambda x:(severity_order.get(x.get("severity","MINOR"),9), SECTIONS.index(x.get("section")) if x.get("section") in SECTIONS else 99))
    score=max(0,100-sum({"CRITICAL":15,"MAJOR":7,"MINOR":2,"EDITORIAL":0}.get(x.get("severity"),1) for x in issues))
    complete=len(stages)==len(SECTIONS)
    if not complete:
        recommendation="BELUM DAPAT DITETAPKAN — REVIEW BELUM LENGKAP"
    elif any(x.get("severity")=="CRITICAL" for x in issues): recommendation="REJECT AND RESUBMIT"
    elif score<65: recommendation="REJECT AND RESUBMIT"
    elif score<80: recommendation="MAJOR REVISION"
    elif score<90: recommendation="MINOR REVISION"
    else: recommendation="ACCEPT"
    return {"summary":f"Review Bertahap telah menghasilkan {len(issues)} temuan dari {len(stages)} bagian yang benar-benar dianalisis. Simpulan diambil dari catatan pada setiap bagian dan tidak mengisi bagian yang belum direview.","issues":issues,"roadmap":roadmap,"score":score,"recommendation":recommendation,"complete":complete,"stage_rows":section_rows,"stages":stages,"research_type":detect_research_type(st.session_state.get("article_text",""))}


def render_staged_results(r):
    st.subheader("📑 Simpulan Catatan Reviewer")
    st.info("Simpulan ini merupakan agregasi langsung dari Review Bertahap. Tidak ada temuan baru yang dibuat pada tahap simpulan.")
    stages=r.get("stages",{}) or {}
    st.write(f"**Bagian yang sudah direview:** {', '.join(x['Bagian'] for x in r.get('stage_rows',[])) or 'belum ada'}")
    if not r.get("complete"):
        st.warning("Review belum lengkap. Rekomendasi editorial final belum ditetapkan; selesaikan Judul sampai Referensi terlebih dahulu.")
    for sec in SECTIONS:
        if sec not in stages:
            st.markdown(f'<div class="summary-section"><h4>{html.escape(sec)}</h4><span class="status-pending">Belum direview</span><br>Belum ada catatan karena bagian ini belum dianalisis.</div>', unsafe_allow_html=True)
            continue
        result=stages[sec]
        fs=result.get("findings",[]) or []
        st.markdown(f'<div class="summary-section"><h4>{html.escape(sec)}</h4><span class="status-reviewed">✓ Sudah direview · {result.get("words",0):,} kata · {len(fs)} catatan</span></div>', unsafe_allow_html=True)
        for i,f in enumerate(fs,1):
            if not isinstance(f,dict):
                st.markdown(f"**{i}.** {f}"); continue
            sev=f.get("severity","MINOR")
            st.markdown(f"**{i}. {sev} — {f.get('issue','')}**")
            st.markdown(f'<div class="evidence-box"><b>📌 Dasar/Bukti pemeriksaan</b><br>{html.escape(str(f.get("evidence","Tidak tersedia.")))}</div>',unsafe_allow_html=True)
            rec=f.get("reviewer_recommendation") or f.get("action") or "Perjelas temuan berdasarkan bukti naskah."
            st.markdown(f'<div class="recommendation-box"><b>Rekomendasi reviewer:</b><br>{html.escape(str(rec))}</div>',unsafe_allow_html=True)
            sols=f.get("alternatives") or professional_alternatives(f,sec,result.get("text",""))
            st.markdown("**Alternatif solusi**")
            for j,sol in enumerate(sols,1):
                st.markdown(f'<div class="solution-box"><b>{chr(64+j)}.</b> {html.escape(sol)}</div>',unsafe_allow_html=True)
    st.markdown("### Kesimpulan umum reviewer")
    st.markdown(f"**Rekomendasi editorial sementara:** {r.get('recommendation','-')} · **Skor indikatif:** {r.get('score','-')}/100")
    st.write("Keputusan editorial final sebaiknya dibuat setelah Judul–Referensi selesai direview dan konsistensi antarbagiannya diperiksa.")

def render_review(r):
    st.subheader("📊 Ringkasan Review")
    c1,c2,c3=st.columns(3); c1.metric("Skor",f"{r.get('score','-')}/100"); c2.metric("Rekomendasi",r.get("recommendation","-")); c3.metric("Prioritas","Substansi → Metode → Gap/Novelty")
    st.markdown(f"**Ringkasan:** {r.get('summary','')}")
    if r.get("research_type"): st.info(f"**Jenis penelitian teridentifikasi:** {r.get('research_type')}")
    tabs=st.tabs(["Gap & Novelty","Metodologi","Konsistensi","Issues","Roadmap"])
    with tabs[0]:
        st.markdown("### Research Gap"); st.write(r.get("gap","")); st.markdown("### Novelty"); st.write(r.get("novelty",""))
    with tabs[1]: st.write(r.get("method",""))
    with tabs[2]:
        st.write(r.get("consistency",""))
        metrics=r.get("metrics",{}).get("section_scores",{}) or r.get("metrics",{}).get("score_domains",{})
        if metrics:
            st.markdown("### Skor per domain"); st.dataframe([{"Domain":k,"Skor":v} for k,v in metrics.items()],use_container_width=True,hide_index=True)
    with tabs[3]:
        issues=r.get("issues",[])
        if not issues: st.success("Tidak ada issue yang terdeteksi pada diagnosis awal.")
        for x in issues:
            if isinstance(x,dict):
                sev=x.get("severity","MINOR"); section=x.get("section",""); msg=x.get("issue",""); evidence=x.get("evidence",""); action=x.get("action","")
                cls=sev.lower() if sev.lower() in ["critical","major","minor","editorial"] else "minor"
                st.markdown(f'<div class="issue-{cls}"><b>{sev}</b> · {section}<br>{msg}<br><span class="small"><b>Bukti:</b> {evidence}<br><b>Tindakan:</b> {action}</span></div>',unsafe_allow_html=True)
            else: st.markdown(f"**{x[0]}** — {x[1]}")
    with tabs[4]:
        for i,x in enumerate(r.get("roadmap",[]),1): st.markdown(f"**{i}.** {x}")
    if r.get("metrics"):
        st.markdown("### Statistik pemeriksaan"); m=r["metrics"]; st.write(f"Kata: **{m.get('words','-'):,}** · Indikator sitasi: **{m.get('citations','-')}** · Entri referensi terdeteksi: **{m.get('reference_lines','-')}**")


def export_markdown(r):
    out="# AI ARTICLE REVIEWER\n\n"
    out+=f"**File:** {st.session_state.filename}\n\n**Score:** {r.get('score')}/100\n\n**Recommendation:** {r.get('recommendation')}\n\n**Research type:** {r.get('research_type','-')}\n\n"
    for title,key in [("Summary","summary"),("Research Gap","gap"),("Novelty","novelty"),("Methodology","method"),("Internal Consistency","consistency")]: out+=f"## {title}\n{r.get(key,'')}\n\n"
    out+="## Issues\n"
    for x in r.get("issues",[]):
        if isinstance(x,dict): out+=f"- **{x.get('severity','MINOR')}** · {x.get('section','')} — {x.get('issue','')}\n  - Evidence: {x.get('evidence','')}\n  - Action: {x.get('action','')}\n"
        else: out+=f"- **{x[0]}** — {x[1]}\n"
    out+="\n## Revision Roadmap\n"+"\n".join(f"{i+1}. {x}" for i,x in enumerate(r.get("roadmap",[])))
    return out

st.markdown('<div class="hero"><h1>📝 AI Article Reviewer</h1><div>Review artikel ilmiah secara sistematis, objektif, dan terarah.</div></div>',unsafe_allow_html=True)


def generate_scope_alternatives(title, field):
    """Generate five usable research-scope alternatives directly from the active title."""
    source=normalize_space(title or "")
    fld=normalize_space(field or "")
    if not source and not fld: return []
    base=source or fld
    stop={"penelitian","research","study","studi","analisis","analysis","pengaruh","hubungan","peran","implementasi","evaluasi","dalam","untuk","terhadap","the","and","of","dan","pada","with","based","kajian"}
    terms=[w for w in re.findall(r"[A-Za-zÀ-ÿ][A-Za-zÀ-ÿ\-]{3,}",base.lower()) if w not in stop]
    topic=" ".join(terms[:9]) if terms else base
    field_label=fld or "bidang keilmuan yang tampak pada judul"
    return [
        {"label":"Scope Fokus Utama","text":f"Membatasi penelitian pada {topic}. Fokus ini mempertahankan persoalan utama dalam judul tanpa memperluas objek di luar naskah."},
        {"label":"Scope Objek/Unit Analisis","text":f"Membatasi unit analisis pada objek, subjek, organisasi, teks, atau fenomena yang benar-benar digunakan dalam penelitian {topic}."},
        {"label":"Scope Konseptual/Teoretis","text":f"Membatasi telaah pada konsep atau teori yang digunakan untuk menjelaskan {topic}; teori tambahan hanya dimasukkan jika diperlukan untuk menjawab pertanyaan penelitian."},
        {"label":"Scope Empiris","text":f"Membatasi ruang empiris berdasarkan lokasi, populasi/informan, periode, atau kasus yang benar-benar tersedia dalam artikel tentang {topic}."},
        {"label":"Scope Kontribusi","text":f"Membatasi klaim kontribusi pada temuan yang dapat dipertanggungjawabkan dari penelitian {topic}, terutama pada ranah {field_label}."},
    ]

def render_scope_cards(scopes):
    if not scopes: return
    st.markdown("### 🎯 Alternatif Scope yang Dapat Dipilih")
    st.caption("Scope dihasilkan dari judul/bidang aktif. Pilih yang paling sesuai dengan isi naskah; sistem tidak menganggap alternatif ini sebagai fakta penelitian.")
    for i,sc in enumerate(scopes,1):
        st.markdown(f'<div class="scope-card"><b>{i}. {html.escape(sc["label"])}</b><br>{html.escape(sc["text"])}</div>',unsafe_allow_html=True)


with st.sidebar:
    _profile_path=Path(__file__).with_name("profile.jpg")
    if _profile_path.exists():
        import base64
        _profile_b64=base64.b64encode(_profile_path.read_bytes()).decode("ascii")
        st.markdown(f"<div class='sidebar-profile'><img src='data:image/jpeg;base64,{_profile_b64}' alt='Foto profil'><div class='name'>AI Article Reviewer</div><div class='role'>Review artikel ilmiah secara sistematis, objektif, dan terarah</div><div class='line'></div></div>", unsafe_allow_html=True)
    st.header("⚙️ Pengaturan")
    mode=st.radio("Mode AI",["Demo / Offline","AI API"])
    api_key=""; endpoint="https://api.openai.com/v1"; model="gpt-4o-mini"
    if mode=="AI API":
        api_key=st.text_input("API Key",type="password")
        endpoint=st.text_input("Endpoint OpenAI-compatible",endpoint)
        model=st.text_input("Model",model)
        st.caption("AI API menghasilkan review substantif. Kunci API tidak disimpan ke file.")
    st.divider(); st.caption("Versi 2.9 — Evidence-grounded review, rekomendasi spesifik, alternatif solusi berbasis temuan, scope otomatis, judul utuh, dan tabel tanpa geser.")

menu=st.radio("Menu",["📄 Artikel Baru","🔍 Review Bertahap","📝 Simpulan Catatan Reviewer","🗂️ Riwayat"],horizontal=True)

if menu=="📄 Artikel Baru":
    st.header("Artikel Baru")
    uploaded=st.file_uploader("Upload artikel",type=["pdf","docx"])
    if uploaded:
        try:
            data=uploaded.read()
            if uploaded.name.lower().endswith(".pdf"):
                text, extract_meta = extract_pdf(data)
            else:
                text, extract_meta = extract_docx(data)
            st.session_state.article_text=text
            st.session_state.filename=uploaded.name
            st.session_state.extract_meta=extract_meta
            st.success(f"Artikel berhasil dibaca UTUH: {len(text.split()):,} kata · {len(text):,} karakter.")
        except Exception as e: st.error(f"Gagal membaca dokumen: {e}")
    if st.session_state.article_text:
        m=st.session_state.get("extract_meta",{})
        c1,c2,c3,c4=st.columns(4)
        c1.metric("Kata terbaca",f"{len(st.session_state.article_text.split()):,}")
        c2.metric("Karakter terbaca",f"{len(st.session_state.article_text):,}")
        c3.metric("Halaman PDF",m.get("pages","-") if m.get("pages") is not None else "-")
        c4.metric("Bagian terdeteksi",sum(bool(v) for k,v in split_sections(st.session_state.article_text).items() if k!="References"))
        if m.get("empty_pages"):
            st.warning("Halaman berikut tidak menghasilkan teks saat ekstraksi PDF: " + ", ".join(map(str,m["empty_pages"])) + ". Naskah perlu diperiksa karena halaman berupa scan/gambar mungkin memerlukan OCR.")
        st.markdown("### 📖 Naskah Lengkap yang Dibaca Sistem")
        st.caption("Seluruh teks hasil ekstraksi ditampilkan dari awal sampai akhir. Tidak ada pembatasan 10.000 karakter.")
        st.text_area("Naskah artikel lengkap",st.session_state.article_text,height=650,key="full_article_view")
        st.download_button("⬇️ Simpan Teks Hasil Ekstraksi",st.session_state.article_text,file_name="naskah_hasil_ekstraksi.txt",mime="text/plain")
        sec_preview=split_sections(st.session_state.article_text)
        detected_title=sec_preview.get("Judul","")
        st.markdown("### 🏷️ Judul artikel yang dibaca sistem")
        if detected_title:
            st.markdown(f'<div class="title-full">{html.escape(detected_title)}</div>',unsafe_allow_html=True)
        else:
            st.warning("Judul belum berhasil diidentifikasi secara utuh. Anda dapat memasukkannya manual.")
        manual_title=st.text_input("Koreksi judul jika ekstraksi belum tepat",detected_title)
        c1,c2=st.columns(2)
        with c1:
            field=st.text_input("Bidang keilmuan","")
            target=st.text_input("Target jurnal (opsional)","")
            auto_scopes=generate_scope_alternatives(manual_title,field)
            render_scope_cards(auto_scopes)
        with c2:
            lang=st.selectbox("Bahasa review",["Indonesia","English"])
            rtype=st.selectbox("Mode",["Full Review","Reviewer Akademik","Reviewer Metodologi","Reviewer Editor"])
        if st.button("🚀 Mulai Full Review",use_container_width=True):
            meta={"field":field,"target_journal":target,"language":lang,"mode":rtype,"active_title":manual_title}
            try:
                if mode=="AI API" and api_key.strip(): result=ai_review(st.session_state.article_text,meta,api_key,endpoint,model)
                else: result=demo_review(st.session_state.article_text,meta)
                result=validate_review_evidence(result,st.session_state.article_text)
                st.session_state.review=result
                st.session_state.history.insert(0,{"file":st.session_state.filename,"time":datetime.now().strftime("%Y-%m-%d %H:%M"),"score":result.get("score"),"recommendation":result.get("recommendation")})
                st.success("Review selesai. Untuk simpulan berbasis per bagian, gunakan menu Simpulan Catatan Reviewer setelah Review Bertahap selesai.")
            except Exception as e: st.error(f"Review gagal: {e}")

elif menu=="🔍 Review Bertahap":
    st.header("Review Bertahap")
    if not st.session_state.article_text: st.info("Upload artikel terlebih dahulu.")
    else:
        sec=split_sections(st.session_state.article_text); selected=st.selectbox("Bagian yang direview",SECTIONS); text=sec.get(selected,"")
        if selected=="References":
            st.write(f"**{count_reference_entries(text):,} entri referensi terdeteksi · {len(text.split()):,} kata**")
            reference_input=st.text_area("Masukan referensi",text,height=360,help="Teks ini dapat diperbaiki secara manual bila ekstraksi referensi dari PDF/DOCX belum sempurna.")
            text=reference_input
        else:
            st.write(f"**{len(text.split()):,} kata terdeteksi**")
            text=st.text_area("Teks bagian",text,height=300)
        if st.button("Analisis bagian ini"):
            if not text: st.warning("Bagian belum berhasil diekstraksi. Periksa heading dokumen atau masukkan bagian secara manual.")
            else:
                result=analyze_section_demo(selected,text)
                # Store every staged review so Hasil Review can use exactly the same findings.
                st.session_state.stage_reviews[selected]=result
                st.session_state.stage_text[selected]=text
                render_section_analysis(result)
                if selected=="References" and result.get("rows"):
                    st.markdown("### Audit Referensi")
                    render_reference_table(result["rows"])
                st.success(f"Review {selected} tersimpan dan sudah terhubung ke Simpulan Catatan Reviewer.")

elif menu=="📝 Simpulan Catatan Reviewer":
    st.header("Simpulan Catatan Reviewer")
    staged=build_staged_review()
    if not staged:
        st.info("Belum ada hasil Review Bertahap. Silakan review bagian artikel secara bertahap mulai Judul sampai Referensi.")
    else:
        render_staged_results(staged)


else:
    st.header("Riwayat Review")
    if not st.session_state.history: st.info("Belum ada riwayat.")
    else: st.dataframe(st.session_state.history,use_container_width=True)
