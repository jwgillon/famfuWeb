import io
import logging
import os
import traceback
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

app.mount("/static", StaticFiles(directory="static"), name="static")

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
        if shape.name == shape_name:
            if shape.has_text_frame:
                para = shape.text_frame.paragraphs[0]
                if not para.runs:
                    para.add_run().text = text
                else:
                    para.runs[0].text = text
                    for run in para.runs[1:]:
                        run.text = ""
                return True
        # Search inside groups
        if shape.shape_type == 6 and hasattr(shape, 'shapes'):
            for sub in shape.shapes:
                if sub.name == shape_name and sub.has_text_frame:
                    para = sub.text_frame.paragraphs[0]
                    if not para.runs:
                        para.add_run().text = text
                    else:
                        para.runs[0].text = text
                        for run in para.runs[1:]:
                            run.text = ""
                    return True
    return False


def write_question(prs, q_num: int, question: str, answers: list, points: list):
    n = q_num

    # Data slide (index 40 = slide 41)
    data_slide = prs.slides[40]
    set_shape_text(data_slide, f"FF_Q{n}", question)
    for i in range(8):
        set_shape_text(data_slide, f"FF_Q{n}_A{i+1}", str(answers[i]))
        set_shape_text(data_slide, f"FF_Q{n}_P{i+1}", str(points[i]))

    # Intro slide (FF_QTN)
    for slide in prs.slides:
        if set_shape_text(slide, f"FF_QT{n}", question):
            break

    # Game board slide (FF_QQN) — answers/points live in groups here
    for slide in prs.slides:
        if set_shape_text(slide, f"FF_QQ{n}", question):
            for i in range(8):
                set_shape_text(slide, f"FF_Q{n}_AT{i+1}", str(answers[i]))
                set_shape_text(slide, f"FF_Q{n}_AP{i+1}", str(points[i]))
            break

    # Reveal slide (FF_QTRN)
    for slide in prs.slides:
        if set_shape_text(slide, f"FF_QTR{n}", question):
            break

    log.info("Q%d: '%s'", n, question[:60])


@app.post("/generate")
async def generate(req: GenerateRequest):
    data = req.game_data

    if "questions" not in data:
        raise HTTPException(400, "Missing field: questions")

    questions = data["questions"]
    if len(questions) != 10:
        raise HTTPException(400, f"Expected 10 questions, got {len(questions)}")

    for i, q in enumerate(questions):
        if len(q.get("answers", [])) != 8 or len(q.get("points", [])) != 8:
            raise HTTPException(400, f"Question {i+1} must have 8 answers and 8 points")

    if not os.path.exists(TEMPLATE_PATH):
        raise HTTPException(500, f"Template not found: {TEMPLATE_PATH}")

    try:
        prs = Presentation(TEMPLATE_PATH)
        log.info("Template loaded. Slides: %d", len(prs.slides))

        for i, q in enumerate(questions):
            write_question(
                prs,
                q_num=i + 1,
                question=str(q["q"]),
                answers=[str(a) for a in q["answers"]],
                points=[str(p) for p in q["points"]],
            )

        slug = "".join(c for c in str(req.theme)[:20] if c.isalnum() or c in " -_").strip() or "Game"

        buf = io.BytesIO()
        prs.save(buf)
        file_bytes = buf.getvalue()
        log.info("File built: %d bytes", len(file_bytes))

        return Response(
            content=file_bytes,
            media_type="application/vnd.ms-powerpoint.presentation.macroEnabled.12",
            headers={"Content-Disposition": f'attachment; filename="Family Feud - {slug}.pptm"'},
        )

    except HTTPException:
        raise
    except Exception as e:
        log.error("Generation failed:\n%s", traceback.format_exc())
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
