import io
import logging
import os
import traceback
from datetime import datetime
from typing import Any

from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from pptx import Presentation
from pydantic import BaseModel

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger(__name__)

app = FastAPI(title="Family Feud Generator API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["POST", "GET"],
    allow_headers=["*"],
)

HERE = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_PATH = os.path.join(HERE, "template.pptm")


class GenerateRequest(BaseModel):
    game_data: dict[str, Any]
    theme: str = "Game"


def set_shape_text(slide, shape_name: str, text: str) -> bool:
    for shape in slide.shapes:
        if shape.name != shape_name:
            continue
        if not shape.has_text_frame:
            log.warning("Shape %s has no text frame", shape_name)
            return False
        tf = shape.text_frame
        if not tf.paragraphs:
            return False
        para = tf.paragraphs[0]
        if not para.runs:
            para.add_run().text = text
        else:
            para.runs[0].text = text
            for run in para.runs[1:]:
                run.text = ""
        return True
    log.warning("Shape not found: %s", shape_name)
    return False


def set_shape_text_all_slides(prs, shape_name: str, text: str) -> bool:
    for slide in prs.slides:
        if set_shape_text(slide, shape_name, text):
            return True
    return False


def format_question_block(n: int, question: str, answers: list, points: list) -> str:
    """Format a question block for the print answer sheet."""
    lines = [f"QUESTION {n}: {question.upper()}"]
    for i, (ans, pts) in enumerate(zip(answers, points), 1):
        prefix = f"{i}) {ans}"
        padded = prefix.ljust(20, '.')
        lines.append(f"{padded}({pts})")
    return "\n".join(lines)


@app.post("/generate")
async def generate(req: GenerateRequest):
    data = req.game_data

    if "questions" not in data:
        raise HTTPException(400, "Missing field: questions")

    questions = data["questions"]
    if len(questions) != 10:
        raise HTTPException(400, f"Expected 10 questions, got {len(questions)}")

    for i, q in enumerate(questions):
        if "q" not in q or "answers" not in q or "points" not in q:
            raise HTTPException(400, f"Question {i+1} missing q/answers/points")
        if len(q["answers"]) != 8:
            raise HTTPException(400, f"Question {i+1} needs 8 answers, got {len(q['answers'])}")
        if len(q["points"]) != 8:
            raise HTTPException(400, f"Question {i+1} needs 8 points, got {len(q['points'])}")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(500, f"Template not found: {TEMPLATE_PATH}")

    log.info("Generating Family Feud. Theme: %s", req.theme)

    try:
        prs = Presentation(TEMPLATE_PATH)
        log.info("Template loaded. Slides: %d", len(prs.slides))

        data_slide  = prs.slides[40]   # Slide 41 - data
        print_slide = prs.slides[2]    # Slide 3  - print answer sheet

        # Build formatted blocks for all 10 questions
        blocks = []
        for q_idx, q in enumerate(questions):
            n = q_idx + 1
            question_text = str(q["q"])
            answers = q["answers"]
            points  = [str(p) for p in q["points"]]

            log.info("Q%d: '%s'", n, question_text[:50])

            # ── Data slide (slide 41) ──
            set_shape_text(data_slide, f"FF_Q{n}", question_text)
            for a_idx in range(8):
                set_shape_text(data_slide, f"FF_Q{n}_A{a_idx+1}", str(answers[a_idx]))
                set_shape_text(data_slide, f"FF_Q{n}_P{a_idx+1}", str(points[a_idx]))

            # ── Game slides — question text on all 3 slides per question ──
            for shape_name in [f"FF_QT{n}", f"FF_QQ{n}", f"FF_QTR{n}"]:
                set_shape_text_all_slides(prs, shape_name, question_text)

            # ── Answer board slide (answers + points) ──
            board_slide_idx = 4 + (q_idx * 3) + 1
            if board_slide_idx < len(prs.slides):
                board_slide = prs.slides[board_slide_idx]
                for a_idx in range(8):
                    set_shape_text(board_slide, f"FF_Q{n}_A{a_idx+1}", str(answers[a_idx]))
                    set_shape_text(board_slide, f"FF_Q{n}_P{a_idx+1}", str(points[a_idx]))

            # ── Reveal slide ──
            reveal_slide_idx = 4 + (q_idx * 3) + 2
            if reveal_slide_idx < len(prs.slides):
                reveal_slide = prs.slides[reveal_slide_idx]
                for a_idx in range(8):
                    set_shape_text(reveal_slide, f"FF_Q{n}_A{a_idx+1}", str(answers[a_idx]))
                    set_shape_text(reveal_slide, f"FF_Q{n}_P{a_idx+1}", str(points[a_idx]))

            blocks.append(format_question_block(n, question_text, answers, points))

        # ── Print answer sheet (slide 3) ──
        # Column layout matches original: Q1-2 | Q3-5 | Q6-8 | Q9-10
        col1 = "\n\n".join(blocks[0:2])
        col2 = "\n\n".join(blocks[2:5])
        col3 = "\n\n".join(blocks[5:8])
        col4 = "\n\n".join(blocks[8:10])

        set_shape_text(print_slide, "PrintColumn1", col1)
        set_shape_text(print_slide, "PrintColumn2", col2)
        set_shape_text(print_slide, "PrintColumn3", col3)
        set_shape_text(print_slide, "PrintColumn4", col4)

        now = datetime.now()
        date_str = now.strftime("%d %B %Y")
        time_str = now.strftime("%I:%M %p")
        print_info = f"FAMILY FEUD:\nHOST ANSWER KEY\nDate: {date_str}\nTime: {time_str}"
        set_shape_text(print_slide, "PrintInfo", print_info)

        log.info("Print answer sheet updated")

        # ── Build file ──
        slug = "".join(c for c in req.theme[:20] if c.isalnum() or c in " -_").strip() or "Game"

        buf = io.BytesIO()
        prs.save(buf)
        buf.seek(0)
        file_bytes = buf.read()
        log.info("File built: %d bytes", len(file_bytes))

        return Response(
            content=file_bytes,
            media_type="application/vnd.ms-powerpoint.presentation.macroEnabled.12",
            headers={"Content-Disposition": f'attachment; filename="Family Feud - {slug}.pptm"'},
        )

    except HTTPException:
        raise
    except Exception as e:
        log.error("Generation failed: %s", traceback.format_exc())
        raise HTTPException(500, f"Failed to generate file: {str(e)}")


# Serve static files
static_dir = os.path.join(HERE, "static")
if os.path.exists(static_dir):
    app.mount("/static", StaticFiles(directory=static_dir), name="static")


@app.get("/", response_class=HTMLResponse)
async def root():
    html_path = os.path.join(HERE, "index.html")
    if not os.path.exists(html_path):
        raise HTTPException(404, "index.html not found")
    with open(html_path) as f:
        return f.read()
