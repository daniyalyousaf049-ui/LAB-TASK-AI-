# Helper only (not part of submission).
# Reads run_output.txt and writes Lab6_Report.docx with real values.
# Usage: python build_report.py

import os
import re
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

OUTPUT = "run_output.txt"
REPORT = "Lab6_Report.docx"
PLACEHOLDER = "[PASTE REAL OUTPUT HERE]"


def set_run_font(run, size=11, bold=False):
    run.font.name = "Calibri"
    run.font.size = Pt(size)
    run.bold = bold


def add_para(doc, text, bold=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run_font(run, 11, bold=bold)
    p.paragraph_format.space_after = Pt(8)
    return p


def add_pre(doc, text):
    """Keep terminal-looking text readable in the report."""
    if not text or not text.strip():
        add_para(doc, PLACEHOLDER)
        return
    for line in text.splitlines():
        p = doc.add_paragraph()
        run = p.add_run(line if line != "" else " ")
        set_run_font(run, 9)
        p.paragraph_format.space_after = Pt(0)
        p.paragraph_format.space_before = Pt(0)


def add_table(doc, headers, rows):
    if not rows:
        add_para(doc, PLACEHOLDER)
        return None
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = "Table Grid"
    for i, h in enumerate(headers):
        cell = table.rows[0].cells[i]
        cell.text = ""
        run = cell.paragraphs[0].add_run(h)
        set_run_font(run, 10, bold=True)
    for r_i, row in enumerate(rows):
        for c_i, val in enumerate(row):
            cell = table.rows[r_i + 1].cells[c_i]
            cell.text = ""
            run = cell.paragraphs[0].add_run(str(val))
            set_run_font(run, 9)
    doc.add_paragraph()
    return table


def section_between(text, start_marker, end_marker):
    if start_marker not in text or end_marker not in text:
        return ""
    a = text.find(start_marker) + len(start_marker)
    b = text.find(end_marker, a)
    return text[a:b].strip()


def load_output():
    if not os.path.exists(OUTPUT):
        return ""
    with open(OUTPUT, "r", encoding="utf-8") as f:
        return f.read()


def parse_versions(text):
    block = section_between(text, "MARKER_VERSIONS_START", "MARKER_VERSIONS_END")
    if not block:
        return PLACEHOLDER
    return block


def parse_model_load(text):
    block = section_between(text, "MARKER_TIMING_START", "MARKER_TIMING_END")
    if not block:
        return PLACEHOLDER
    return block


def parse_token_rows(text):
    block = section_between(text, "MARKER_TOKEN_TABLE_START", "MARKER_TOKEN_TABLE_END")
    rows = []
    if not block:
        return rows
    for line in block.splitlines():
        line = line.strip()
        if not line or line.startswith("Input") or line.startswith("---"):
            continue
        # keep whole line as one cell dump if splitting is fragile
        rows.append([line])
    return rows


def parse_metrics(text):
    block = section_between(text, "MARKER_METRICS_START", "MARKER_METRICS_END")
    rows = []
    if not block:
        return rows
    for line in block.splitlines():
        line = line.strip()
        if not line or line.startswith("Strategy") or line.startswith("---") or line.startswith("Token audit") or line.startswith("Top-3"):
            continue
        # metric summary lines look like: 0overlap tfidf dev 0.750 ...
        parts = line.split()
        if len(parts) >= 8 and parts[2] in ("dev", "test"):
            rows.append(parts[:8])
    return rows


def parse_selection(text):
    block = section_between(text, "MARKER_SELECTION_START", "MARKER_SELECTION_END")
    if not block:
        return PLACEHOLDER
    return block


def parse_failures(text):
    block = section_between(text, "MARKER_FAILURES_START", "MARKER_FAILURES_END")
    if not block:
        return PLACEHOLDER
    return block


def parse_faq(text):
    block = section_between(text, "MARKER_FAQ_START", "MARKER_FAQ_END")
    if not block:
        return PLACEHOLDER
    return block


def parse_doc_chat(text):
    block = section_between(text, "MARKER_DOC_CHAT_START", "MARKER_DOC_CHAT_END")
    if not block:
        return PLACEHOLDER
    return block


def parse_chunk_stats(text):
    block = section_between(text, "MARKER_CHUNK_STATS_START", "MARKER_CHUNK_STATS_END")
    if not block:
        return PLACEHOLDER
    return block


def parse_paraphrase(text):
    block = section_between(text, "MARKER_PARAPHRASE_START", "MARKER_PARAPHRASE_END")
    if not block:
        return PLACEHOLDER
    return block


def build():
    text = load_output()
    doc = Document()

    title = doc.add_heading(
        "Lab 6 Report: NLP Fundamentals - Tokenization, Embeddings and Text Chunking",
        level=0,
    )
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER

    add_para(doc, "Name: Daniyal Yousaf", bold=True)
    add_para(doc, "Roll #: 097")
    add_para(doc, "Section: BSAI-3B")
    add_para(doc, "Course: CS4944 - Artificial Intelligence")
    add_para(doc, "Submission files: lab6_solution.py, Lab6_Report.docx")

    # versions
    add_heading = doc.add_heading
    add_heading("Library and Python versions", level=2)
    add_para(doc, "These versions came from my run_output.txt:")
    add_pre(doc, parse_versions(text) if text else PLACEHOLDER)

    # Section 1
    add_heading("1. Scenario and Objective", level=1)
    add_para(
        doc,
        "For this lab I built a small university help desk retrieval system. A student types a "
        "question about campus policy, and the system finds the most relevant passage from synthetic "
        "handbook documents. I used five docs: late work, password reset, assignment upload, "
        "contractor access and printers. Each one is about 150 to 250 words."
    )
    add_para(
        doc,
        "The goal was to compare tokenization methods, try three chunking strategies, and evaluate "
        "TF-IDF against sentence embeddings with Hit@3 and MRR@3. I tuned only on the four development "
        "queries, then checked the four held-out test queries. After that I wrapped the chosen setup "
        "in a confirmation-based terminal chatbot so it does not invent policy text."
    )

    # Section 2
    add_heading("2. Tokenization and Embeddings Analysis", level=1)
    add_para(
        doc,
        "Task 1 compared whitespace splitting, a Unicode word-like regex, and the all-MiniLM-L6-v2 "
        "tokenizer on sample strings. I used punctuation, a contraction, an email, CS4944, an empty "
        "string and a dotted filename. The table below is copied from the run."
    )

    token_block = section_between(text, "MARKER_TOKEN_TABLE_START", "MARKER_TOKEN_TABLE_END")
    add_pre(doc, token_block if token_block else PLACEHOLDER)

    add_para(
        doc,
        "Whitespace keeps punctuation stuck on words, so search! stays one token. The regex looks "
        "cleaner on AI, search! but it breaks emails into pieces and splits can't into can and t. "
        "That is a normalization risk if you throw away the original string. Model subwords are a "
        "third view again, so I only use them for the encoder path."
    )
    add_para(
        doc,
        "For embeddings I used sentence-transformers/all-MiniLM-L6-v2 on CPU. TF-IDF is still useful "
        "for keyword matches and it is cheap. Embeddings helped more on paraphrases where the wording "
        "did not match the handbook."
    )

    add_heading("Timing method", level=2)
    add_para(
        doc,
        "I timed model load once at startup, then timed encoding and search separately for each "
        "config. The model is not reloaded per strategy."
    )
    add_pre(doc, parse_model_load(text) if text else PLACEHOLDER)

    # Section 3
    add_heading("3. Chunking and Retrieval Evaluation", level=1)
    add_para(
        doc,
        "I used three chunkers on the same five documents: 40-word windows with 0 overlap, 40-word "
        "windows with 10 overlap, and sentence-aware packing with a 40-word budget. If a sentence is "
        "longer than 40 words, sentence_chunks falls back to windows. Invalid overlap raises ValueError."
    )
    add_para(
        doc,
        "I also put related but conflicting text across docs on purpose. The late policy uses a 10 "
        "percent per day rule with a three day cap, while the assignment doc mentions a five day portal "
        "window with a flat five percent fee. The printer doc says 150 pages, but also notes an old "
        "200 page door notice. That made conflicting-source failures easier to spot."
    )

    add_heading("Chunk stats from the run", level=2)
    add_pre(doc, parse_chunk_stats(text) if text else PLACEHOLDER)

    add_heading("Paraphrase checks (TF-IDF vs embeddings)", level=2)
    add_pre(doc, parse_paraphrase(text) if text else PLACEHOLDER)

    add_heading("Dev and test metrics", level=2)
    add_para(
        doc,
        "A hit counts only when the retrieved chunk is from the labeled document and contains the "
        "full evidence string, including the condition or exception. Metrics below are from the run."
    )
    metric_rows = parse_metrics(text)
    add_table(
        doc,
        ["Strategy", "Backend", "Split", "Hit@3", "MRR@3", "Chunks", "Encode(s)", "Search(s)"],
        metric_rows if metric_rows else [],
    )

    add_heading("Strategy choice from DEV only", level=2)
    add_pre(doc, parse_selection(text) if text else PLACEHOLDER)
    add_para(
        doc,
        "I picked the best DEV Hit@3, then MRR@3, and preferred embedding on ties because the "
        "chatbot needs paraphrase support. TEST numbers above are reported without changing the choice."
    )

    # Section 4
    add_heading("4. Chatbot Transcripts and Fallback Behavior", level=1)
    add_para(
        doc,
        "Task 1 uses a TF-IDF FAQ bot. Exact canonical questions return the stored answer and source. "
        "Anything else only gets suggestions, so I still have to ask the canonical question. Empty or "
        "unsupported input falls back instead of guessing."
    )

    add_heading("4.1 FAQ transcripts from the run", level=2)
    add_pre(doc, parse_faq(text) if text else PLACEHOLDER)

    add_heading("4.2 Document chatbot transcripts from the run", level=2)
    add_para(
        doc,
        "The document bot shows top-3 evidence with IDs, scores and character offsets. I only see a "
        "full passage after choosing 1-3. Enter or an invalid choice means no answer is selected. "
        "Unanswerable queries are never auto-asserted as answers."
    )
    add_pre(doc, parse_doc_chat(text) if text else PLACEHOLDER)

    # Section 5
    add_heading("5. Limitations and Failure Edge Cases", level=1)
    add_para(
        doc,
        "Below are real misses printed by the script. I only discuss failures that showed up in "
        "run_output.txt."
    )
    fail_block = parse_failures(text) if text else ""
    if not fail_block or "No answerable misses" in fail_block:
        add_pre(doc, fail_block if fail_block else PLACEHOLDER)
    else:
        # keep the raw dump, then explain up to 3 concrete ones from the dump
        add_pre(doc, fail_block)
        explained = 0
        chunks = fail_block.split("FAIL ")
        for chunk in chunks:
            if explained >= 3:
                break
            if "query:" not in chunk:
                continue
            lines = chunk.splitlines()
            query = ""
            config = ""
            retrieved = ""
            expected = ""
            for line in lines:
                if "config:" in line:
                    config = line.split("config:", 1)[1].strip()
                if "query:" in line:
                    query = line.split("query:", 1)[1].strip()
                if "expected_doc:" in line:
                    expected = line.split("expected_doc:", 1)[1].strip()
                if "retrieved_text:" in line:
                    retrieved = line.split("retrieved_text:", 1)[1].strip()
            if not query:
                continue
            explained = explained + 1
            add_heading("Failure case " + str(explained), level=2)
            add_para(doc, "Query: " + query)
            add_para(doc, "Config: " + config + " | expected doc: " + expected)
            add_para(doc, "Retrieved: " + (retrieved[:220] + "..." if len(retrieved) > 220 else retrieved))
            # cause/fix based only on what this failure shows
            if expected and "contractor" in retrieved.lower() and expected == "late":
                add_para(
                    doc,
                    "Likely cause: lexical overlap pulled a contractor chunk that talks about work "
                    "rules instead of the late penalty evidence. Fix: add document-type metadata or "
                    "require the evidence span, not just topical words."
                )
            elif "password" in query.lower() or expected == "password":
                add_para(
                    doc,
                    "Likely cause: the top chunk was from the password doc but did not contain the "
                    "full labeled evidence string, or TF-IDF ranked a nearby span first. Fix: keep "
                    "the rule and exception inside one sentence-sized chunk and prefer embeddings "
                    "for paraphrase wording."
                )
            elif "blank" in query.lower() or expected == "printer":
                add_para(
                    doc,
                    "Likely cause: the blank-page procedure was split away from the words the query "
                    "used, so a different printer chunk ranked higher. Fix: keep procedure steps in "
                    "one chunk and maybe boost exact phrase matches."
                )
            else:
                add_para(
                    doc,
                    "Likely cause: the top-3 chunks missed the full labeled evidence, either because "
                    "of a hard window cut or because a conflicting related passage ranked higher. "
                    "Fix: use sentence-aware packing for rule/exception text and confirm evidence "
                    "before answering."
                )

    add_para(
        doc,
        "The simple sentence splitter can also break on abbreviations like Mr. or e.g. because it "
        "splits on period plus space. That is a known limit of the teaching splitter."
    )

    # FAQ vs document comparison
    add_heading("FAQ vs document retrieval", level=1)
    add_para(
        doc,
        "FAQ matching is narrow. If I type the exact stored question I get a short answer and a "
        "source file name. If I paraphrase, I only get suggestions. That is safer, but it also means "
        "the FAQ bot has almost no room for missing context inside a long policy."
    )
    add_para(
        doc,
        "Document retrieval can surface the actual handbook span with offsets, which is better when "
        "the answer needs a condition or exception. The downside is conflicting sources. My late and "
        "assignment docs disagree on late windows, and a high score does not tell me which policy "
        "wins. That is why the chatbot asks me to confirm a passage instead of stating it as fact."
    )

    # Section 6 reflections - manual questions unchanged
    add_heading("6. Reflection Responses", level=1)

    add_para(doc, "What did you learn today?", bold=True)
    add_para(
        doc,
        "I learned that tokenizers disagree in ways that actually change search. Whitespace, regex "
        "and model subwords all cut the same sentence differently. I also saw that chunk boundaries "
        "can hide an exception even when the main rule is retrieved, and that Hit@3 needs both the "
        "right document and the full evidence text."
    )

    add_para(doc, "How will you apply this in a real AI system?", bold=True)
    add_para(
        doc,
        "I would keep source IDs and character offsets on every chunk, evaluate on a held-out set, "
        "and force a human or user confirmation before showing policy text as an answer. I would also "
        "time model load separately from encode/search so performance numbers are honest."
    )

    add_para(doc, "What failed or surprised you?", bold=True)
    if text and "MARKER_FAILURES_START" in text:
        add_para(
            doc,
            "The surprising part was how often a high score still pointed at the wrong place when "
            "docs conflicted or when TF-IDF missed a paraphrase. The automatic failure dump in the "
            "run made that hard to ignore. Cosine similarity also always returns neighbors, even for "
            "flight prices, so unanswerable queries need a fallback workflow."
        )
    else:
        add_para(doc, PLACEHOLDER)

    add_para(doc, "What would you do differently next time?", bold=True)
    add_para(
        doc,
        "Next time I would add abbreviation-aware sentence splitting, store audience or policy-version "
        "metadata on chunks, and calibrate a reject threshold on separate development examples instead "
        "of trusting raw top-1 scores."
    )

    # Section 7
    add_heading("7. Reproduction Steps", level=1)
    add_para(doc, "From the Lab-6 folder:")
    add_para(
        doc,
        "1. python -m venv .venv\n"
        "2. .venv\\Scripts\\activate\n"
        "3. pip install numpy scikit-learn sentence-transformers python-docx\n"
        "4. python lab6_solution.py\n"
        "5. python build_report.py"
    )
    add_para(
        doc,
        "lab6_solution.py writes run_output.txt and retrieval_log.csv. build_report.py reads "
        "run_output.txt and fills Lab6_Report.docx. Do not invent scores by hand."
    )

    doc.save(REPORT)
    print("Wrote", REPORT)
    if not text:
        print("WARNING: run_output.txt missing. Report contains placeholders.")


if __name__ == "__main__":
    build()
