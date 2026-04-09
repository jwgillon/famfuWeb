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

# VERSION: 1.2 - Testing Answer Space Fix
log.info("--- SYSTEM STARTING: VERSION 1.3 ---")

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


def set_shape_text(shape_collection, shape_name: str, text: str) -> bool:
    """Recursively search shapes (including inside groups) and set text."""
    for shape in shape_collection:
        # Recurse into groups
        if shape.shape_type == 6:  # GROUP
            if set_shape_text(shape.shapes, shape_name, text):
                return True
            continue
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
    return False


def set_slide_shape(slide, shape_name: str, text: str) -> bool:
    return set_shape_text(slide.shapes, shape_name, text)


def set_all_slides(prs, shape_name: str, text: str) -> bool:
    for slide in prs.slides:
        if set_slide_shape(slide, shape_name, text):
            return True
    return False


def format_question_block(n: int, question: str, answers: list, points: list) -> str:
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
            raise HTTPException(400, f"Question {i+1} needs 8 answers")
        if len(q["points"]) != 8:
            raise HTTPException(400, f"Question {i+1} needs 8 points")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(500, f"Template not found: {TEMPLATE_PATH}")

    log.info("Generating Family Feud. Theme: %s", req.theme)

    try:
        prs = Presentation(TEMPLATE_PATH)
        log.info("Template loaded. Slides: %d", len(prs.slides))

        data_slide  = prs.slides[40]  # Slide 41 - data
        print_slide = prs.slides[2]   # Slide 3  - print answer sheet

        blocks = []

        for q_idx, q in enumerate(questions):
            n = q_idx + 1
            question_text = str(q["q"])
            answers = q["answers"]
            points  = [str(p) for p in q["points"]]

            log.info("Q%d: '%s'", n, question_text[:50])

            # ── Data slide (slide 41) ──
            set_slide_shape(data_slide, f"FF_Q{n}", question_text)
            for a_idx in range(8):
                set_slide_shape(data_slide, f"FF_Q{n}_A{a_idx+1}", str(answers[a_idx]))
                set_slide_shape(data_slide, f"FF_Q{n}_P{a_idx+1}", str(points[a_idx]))

            # ── Question text on all 3 game slides for this question ──
            for shape_name in [f"FF_QT{n}", f"FF_QQ{n}", f"FF_QTR{n}"]:
                set_all_slides(prs, shape_name, question_text)

            # ── Answer board slide — answers inside groups (FF_QN_ATX / FF_QN_APX) ──
            board_slide_idx = 4 + (q_idx * 3) + 1
            if board_slide_idx < len(prs.slides):
                board_slide = prs.slides[board_slide_idx]
                for a_idx in range(8):
                    a_num = a_idx + 1
                    ok_a = set_slide_shape(board_slide, f"FF_Q{n}_AT{a_num}", str(answers[a_idx]))
                    ok_p = set_slide_shape(board_slide, f"FF_Q{n}_AP{a_num}", str(points[a_idx]))
                    if not ok_a or not ok_p:
                        log.warning("MISS on board: Q%d A%d  AT=%s AP=%s", n, a_num, ok_a, ok_p)
                    else:
                        log.info("  AT%d='%s' AP%d=%s", a_num, answers[a_idx], a_num, points[a_idx])

            # ── Reveal slide ──
            reveal_slide_idx = 4 + (q_idx * 3) + 2
            if reveal_slide_idx < len(prs.slides):
                reveal_slide = prs.slides[reveal_slide_idx]
                for a_idx in range(8):
                    a_num = a_idx + 1
                    
                    # Capture the success/fail status for each shape
                    ok_ra = set_slide_shape(reveal_slide, f"FF_Q{n}_AT{a_num}_R", str(answers[a_idx]))
                    ok_rp = set_slide_shape(reveal_slide, f"FF_Q{n}_AP{a_num}_R", str(points[a_idx]))
                    
                    # Error logging for the Reveal Slide
                    if not ok_ra or not ok_rp:
                        log.warning("MISS on REVEAL: Q%d A%d  AT=%s AP=%s", n, a_num, ok_ra, ok_rp)
                    else:
                        log.info("  REVEAL SUCCESS: Q%d A%d AT%d='%s' AP%d=%s", n, a_num, a_num, answers[a_idx], a_num, points[a_idx])


            blocks.append(format_question_block(n, question_text, answers, points))

        # ── Print answer sheet (slide 3) ──
        col1 = "\n\n".join(blocks[0:2])
        col2 = "\n\n".join(blocks[2:5])
        col3 = "\n\n".join(blocks[5:8])
        col4 = "\n\n".join(blocks[8:10])

        set_slide_shape(print_slide, "PrintColumn1", col1)
        set_slide_shape(print_slide, "PrintColumn2", col2)
        set_slide_shape(print_slide, "PrintColumn3", col3)
        set_slide_shape(print_slide, "PrintColumn4", col4)

        now = datetime.now()
        print_info = f"FAMILY FEUD:\nHOST ANSWER KEY\nDate: {now.strftime('%d %B %Y')}\nTime: {now.strftime('%I:%M %p')}"
        set_slide_shape(print_slide, "PrintInfo", print_info)
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
