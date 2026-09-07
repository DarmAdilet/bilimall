import os
import json
import uuid
import re
import io
from functools import lru_cache
from datetime import datetime

from django.http import JsonResponse, HttpResponse
from django.views.decorators.csrf import csrf_exempt
from django.shortcuts import render
from django.conf import settings
from django.core.cache import cache

from google import genai
from google.genai import types
from dotenv import load_dotenv

from pptx import Presentation
from pptx.util import Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor

from docx import Document
from docx.shared import Pt as DocxPt, RGBColor as DocxRGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============ CONFIGURATION ============
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise ValueError("⚠️ GEMINI_API_KEY not found in .env!")

client = genai.Client(api_key=API_KEY)
MODEL_NAME = "gemini-2.5-flash"

SAFETY_SETTINGS = [
    types.SafetySetting(category="HARM_CATEGORY_HARASSMENT", threshold="BLOCK_NONE"),
    types.SafetySetting(category="HARM_CATEGORY_HATE_SPEECH", threshold="BLOCK_NONE"),
    types.SafetySetting(category="HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold="BLOCK_NONE"),
    types.SafetySetting(category="HARM_CATEGORY_DANGEROUS_CONTENT", threshold="BLOCK_NONE"),
    types.SafetySetting(category="HARM_CATEGORY_CIVIC_INTEGRITY", threshold="BLOCK_NONE"),
]

CHAT_SYSTEM_INSTRUCTION = """Рөл: Сен мұғалімдерге Информатика пәнінің тақырыптарын басқа жаратылыстану пәндерімен байланыстыруға көмектесетін ассистентсің.

Нұсқаулық:
Мен саған Информатика пәнінен белгілі бір тақырыпты беремін. Сен тек https://okulyk.kz/ сайтындағы қазақша Алгебра, Геометрия, Физика, Химия, Биология, География оқулықтарын қолдана отырып, сол тақырыптың Алгебра, Геометрия, Физика, Химия, Биология, География пәндерінен (5-11 сыныптар) қандай тақырыптармен байланысы бар екенін анықта.

Байланысты тақырыптарды келесі форматта көрсет:
- Пән атауы:
- Сыныбы:
- Оқулықтың авторы және баспасы:
- Тақырыптың атауы:
- Беті:
- Байланысы:

Маңызды:
Тек нақты байланысы бар пәндерді ғана көрсет.
Бір тақырыпты бірнеше рет әртүрлі құрылғылардан берген кезде жауаптар барлық жағдайда бірдей болуы керек, артық немесе әртүрлі жауаптар берме.
Сұрақ тақырыптан ауытқыса немесе басқа салаға қатысты болса, жауап берме.
Жауап бермес бұрын оқулықтарды толық зерттеп, тек содан кейін ғана жауап қайтар.
Always respond in English only. All output must be in English."""

GENERATION_CONFIG = types.GenerateContentConfig(
    temperature=0.3,
    top_p=0.95,
    top_k=40,
    max_output_tokens=8192,
    safety_settings=SAFETY_SETTINGS,
    response_mime_type="text/plain",
    system_instruction=[
        types.Part.from_text(text=CHAT_SYSTEM_INSTRUCTION)
    ],
)


# ============ HELPER FUNCTIONS ============
def get_generation_config(system_instruction=None):
    """Build generation config"""
    config = types.GenerateContentConfig(
        temperature=0.3,
        top_p=0.95,
        top_k=40,
        max_output_tokens=8192,
        safety_settings=SAFETY_SETTINGS,
        response_mime_type="text/plain",
    )
    if system_instruction:
        config.system_instruction = [types.Part.from_text(text=system_instruction)]
    return config


def stream_gemini_response(user_message, system_instruction=None, use_chat_config=False):
    """Get response from Gemini API"""
    if use_chat_config:
        config = GENERATION_CONFIG
    elif system_instruction:
        config = get_generation_config(system_instruction)
    else:
        config = get_generation_config()

    contents = [
        types.Content(
            role="user",
            parts=[types.Part.from_text(text=user_message)]
        )
    ]

    response_text = ""
    try:
        for chunk in client.models.generate_content_stream(
                model=MODEL_NAME,
                contents=contents,
                config=config
        ):
            response_text += chunk.text
    except Exception as e:
        raise Exception(f"Gemini API error: {str(e)}")

    return response_text.strip()


def validate_json_request(request):
    """Validate JSON request"""
    if request.method != 'POST':
        return None, JsonResponse(
            {'error': 'Only POST requests are accepted'},
            status=405
        )

    try:
        if request.content_type == 'application/json':
            data = json.loads(request.body)
        else:
            data = request.POST.dict()
        return data, None
    except json.JSONDecodeError:
        return None, JsonResponse(
            {'error': 'Invalid JSON format'},
            status=400
        )


# ============ CHAT ============
@csrf_exempt
def chat_with_gemini(request):
    """Process user message and return Gemini response."""
    if request.method != "POST":
        return JsonResponse(
            {"error": "⚠️ Only POST requests are supported!"},
            status=400
        )

    try:
        data = json.loads(request.body)
        user_message = data.get("message", "").strip()

        if not user_message:
            return JsonResponse(
                {"response": "⚠️ Message cannot be empty!"},
                status=400
            )

        cache_key = f"chat_{hash(user_message)}"
        cached_response = cache.get(cache_key)

        if cached_response:
            return JsonResponse({"response": cached_response})

        response_text = stream_gemini_response(user_message, use_chat_config=True)

        cache.set(cache_key, response_text, 3600)

        return JsonResponse({"response": response_text})

    except json.JSONDecodeError:
        return JsonResponse(
            {"error": "⚠️ Invalid JSON format!"},
            status=400
        )

    except Exception as e:
        return JsonResponse(
            {"error": f"⚠️ Server error: {str(e)}"},
            status=500
        )


# ============ WORD DOWNLOAD ============
@csrf_exempt
def download_exercise_word(request):
    """Download exercise as Word file"""
    data, error_response = validate_json_request(request)
    if error_response:
        return error_response

    exercise_content = data.get('content', '').strip()
    topic = data.get('topic', 'Exercise').strip()
    level = data.get('level', '').strip()
    format_type = data.get('format', '').strip()

    if not exercise_content:
        return JsonResponse(
            {'error': 'No exercise content'},
            status=400
        )

    try:
        doc = Document()

        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = DocxPt(12)

        title = doc.add_heading(topic, level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title.runs[0]
        title_run.font.size = DocxPt(18)
        title_run.font.bold = True
        title_run.font.color.rgb = DocxRGBColor(0, 84, 149)

        doc.add_paragraph()
        meta_para = doc.add_paragraph()
        meta_para.add_run(f'📊 Grade: ').bold = True
        meta_para.add_run(level)
        meta_para.add_run('\n📝 Format: ').bold = True
        meta_para.add_run(format_type)
        meta_para.add_run('\n📅 Date: ').bold = True
        meta_para.add_run(datetime.now().strftime('%Y-%m-%d'))

        doc.add_paragraph('_' * 60)
        doc.add_paragraph()

        lines = exercise_content.split('\n')
        for line in lines:
            if line.strip():
                if line.strip().isupper() or line.strip().endswith(':'):
                    para = doc.add_heading(line.strip(), level=2)
                    para_run = para.runs[0]
                    para_run.font.size = DocxPt(14)
                    para_run.font.color.rgb = DocxRGBColor(0, 84, 149)
                else:
                    para = doc.add_paragraph(line.strip())
                    para_format = para.paragraph_format
                    para_format.line_spacing = 1.5
                    para_format.space_after = DocxPt(6)

        doc.add_paragraph()
        footer_para = doc.add_paragraph()
        footer_run = footer_para.add_run('\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n')
        footer_run.font.color.rgb = DocxRGBColor(128, 128, 128)
        footer_para.add_run('Generated by: BilimAll AI\n').italic = True
        footer_para.add_run('https://bilimall.kz').font.color.rgb = DocxRGBColor(0, 84, 149)
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        response = HttpResponse(
            doc_io.read(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

        filename = f"{topic.replace(' ', '_')}_exercise.docx"
        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        return response

    except Exception as e:
        print(f"Word error: {str(e)}")
        return JsonResponse({
            'error': f'Error creating Word file: {str(e)}'
        }, status=500)


# ============ GAME FUNCTIONS ============
def generate_game_questions(topic, level):
    """Generate game questions"""
    prompt = f"""Create interactive quiz questions for learning Python programming.

Parameters:
- Topic: {topic}
- Level: {level}

Instructions:
1. Create 10 questions
2. Each question should have 4 answer options
3. Questions should cover Python syntax, functions, loops, conditions, data types
4. Some questions may include code examples
5. Indicate the correct answer index (0-3)
6. Provide a short explanation for each answer

REQUIRED: Return your answer in JSON format only.

JSON format:
{{
  "questions": [
    {{
      "question": "Question text",
      "code": "print('Hello')",
      "options": ["Option A", "Option B", "Option C", "Option D"],
      "correctIndex": 0,
      "explanation": "Explanation"
    }}
  ]
}}

Note: "code" field is optional."""

    system_inst = "You are a Python teaching bot. Return answers in JSON format only!"

    response = stream_gemini_response(prompt, system_inst)

    try:
        response = response.strip()
        if response.startswith('```json'):
            response = response[7:]
        if response.startswith('```'):
            response = response[3:]
        if response.endswith('```'):
            response = response[:-3]
        response = response.strip()

        data = json.loads(response)
        return data.get('questions', [])
    except json.JSONDecodeError as e:
        print(f"JSON parse error: {e}")
        return []


@csrf_exempt
def generate_game_questions_api(request):
    """Game questions generation API"""
    data, error_response = validate_json_request(request)
    if error_response:
        return error_response

    topic = data.get('topic', '').strip()
    level = data.get('level', '').strip()

    if not topic or not level:
        return JsonResponse(
            {'error': 'Topic and level must be specified'},
            status=400
        )

    try:
        cache_key = f"game_{hash(topic + level)}"
        cached_questions = cache.get(cache_key)

        if cached_questions:
            return JsonResponse({
                'success': True,
                'questions': cached_questions
            })

        questions = generate_game_questions(topic, level)

        if not questions:
            return JsonResponse(
                {'error': 'Questions were not generated'},
                status=500
            )

        cache.set(cache_key, questions, 1800)

        return JsonResponse({
            'success': True,
            'questions': questions
        })

    except Exception as e:
        return JsonResponse({
            'error': f'An error occurred: {str(e)}'
        }, status=500)


def game_view(request):
    """Show game page"""
    topic = request.GET.get('topic', 'Python Basics')
    level = request.GET.get('level', 'Grade 5')

    return render(request, 'steam/game.html', {
        'topic': topic,
        'level': level
    })


# ============ EXERCISE GENERATION ============
def generate_exercise_content(topic, format_type, level):
    """Generate exercise content"""
    prompt = f"""Create an exercise based on the following parameters:

Topic: {topic}
Format: {format_type}
Grade level: {level}

Instructions:
1. The exercise must be in Kazakh language
2. Difficulty should match the grade level
3. Create according to the selected format:
   - Test: 20 questions, 4 answer options, correct answers shown
   - Exercise: 5-7 practical tasks with instructions
   - Questions: 8-10 open-ended questions with varying difficulty

4. Tasks should be clear, specific, and aimed at knowledge assessment"""

    system_inst = "You are an experienced teacher. Return only the finished exercise."

    return stream_gemini_response(prompt, system_inst)


@csrf_exempt
def generate_exercise(request):
    """Exercise generation API"""
    data, error_response = validate_json_request(request)
    if error_response:
        return error_response

    topic = data.get('topic', '').strip()
    format_type = data.get('format', '').strip()
    level = data.get('level', '').strip()

    if not topic or not format_type or not level:
        return JsonResponse(
            {'error': 'All fields must be filled'},
            status=400
        )

    if format_type == 'Game':
        from urllib.parse import quote
        return JsonResponse({
            'success': True,
            'redirect': f'/game/?topic={quote(topic)}&level={quote(level)}'
        })

    try:
        cache_key = f"exercise_{hash(topic + format_type + level)}"
        cached_response = cache.get(cache_key)

        if cached_response:
            return JsonResponse({
                'success': True,
                'exercise_content': cached_response
            })

        exercise_content = generate_exercise_content(topic, format_type, level)

        if not exercise_content:
            return JsonResponse(
                {'error': 'Exercise was not generated'},
                status=500
            )

        cache.set(cache_key, exercise_content, 1800)

        return JsonResponse({
            'success': True,
            'exercise_content': exercise_content
        })

    except Exception as e:
        return JsonResponse({
            'error': f'An error occurred: {str(e)}'
        }, status=500)


# ============ SLIDE GENERATION ============
def create_slide_text_frame(shape, content, title_level=False):
    """Format slide text"""
    tf = shape.text_frame
    tf.clear()

    for paragraph in content.split('\n'):
        if not paragraph.strip():
            continue

        p = tf.add_paragraph()
        p.text = paragraph.strip()

        if title_level:
            p.font.size = PptxPt(28)
            p.font.bold = True
            p.font.color.rgb = PptxRGBColor(0, 84, 149)
        else:
            p.font.size = PptxPt(18)
            if any(paragraph.strip().startswith(char) for char in ['•', '-', '1.', '2.', '3.']):
                p.level = 1
                p.font.size = PptxPt(16)

        p.font.name = 'Arial'
        p.space_after = PptxPt(8)


def generate_slide_content(topic):
    """Generate slide content"""
    prompt = f"""The following line contains information about Computer Science and natural science subjects:

{topic}

This line contains: a Computer Science topic, a natural science subject name, grade, and a natural science topic — separated by commas.

Your task:
1. Explain the connection between Computer Science and the natural science subject in the given data.
2. Based on this connection, create a presentation text in Kazakh consisting of 5–7 slides.
3. Each slide should have specific content (e.g.: introduction, key concepts, subject connection, example, conclusion).
4. Use only official textbooks from okulyk.kz.
5. Separate each slide with the label '**N-slide**'.
6. Sentences in each slide should be detailed, clear, and specific.
7. Include a list of references, grade, publisher, and authors."""

    system_inst = "Return only the finished presentation text."
    return stream_gemini_response(prompt, system_inst)


def create_kazakh_slides(prs, topic, content):
    """Create slides"""
    slides_content = re.split(r'\*\*(\d+-slide.*?)\*\*', content)

    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic
    title_slide.placeholders[1].text = "Automatically generated presentation"

    for i in range(1, len(slides_content), 2):
        if i + 1 >= len(slides_content):
            continue

        slide_title = slides_content[i].strip()
        slide_content = slides_content[i + 1].strip()

        if not slide_content:
            continue

        slide = prs.slides.add_slide(prs.slide_layouts[1])
        create_slide_text_frame(slide.shapes.title, slide_title, title_level=True)
        create_slide_text_frame(slide.placeholders[1], slide_content)


@csrf_exempt
def generate_slide(request):
    """Slide generation API"""
    data, error_response = validate_json_request(request)
    if error_response:
        return error_response

    topic = data.get('topic', '').strip()

    if not topic:
        return JsonResponse(
            {'error': 'Topic cannot be empty'},
            status=400
        )

    try:
        slide_text = generate_slide_content(topic)

        if not slide_text:
            return JsonResponse(
                {'error': 'Content was not generated'},
                status=500
            )

        template_path = os.path.join(settings.BASE_DIR, 'steam', 'template.pptx')
        output_filename = f"{uuid.uuid4().hex}.pptx"
        output_path = os.path.join(settings.MEDIA_ROOT, output_filename)

        os.makedirs(settings.MEDIA_ROOT, exist_ok=True)

        prs = Presentation(template_path)
        create_kazakh_slides(prs, topic, slide_text)
        prs.save(output_path)

        return JsonResponse({
            'success': True,
            'slide_text': slide_text,
            'pptx_url': f'/media/{output_filename}'
        })

    except Exception as e:
        return JsonResponse({
            'error': f'An error occurred: {str(e)}',
            'slide_text': 'An error occurred while creating the slide.'
        }, status=500)


# ============ ASSISTANT CHATBOT ============
CHATBOT_SYSTEM_INSTRUCTION = """You are the assistant of the BilimALL AI project.

About BilimALL AI:
- This is an AI platform that connects Computer Science with natural science subjects
- Created at Aktobe Regional University named after K. Zhubanov
- Designed for grades 5-11

Services:
1. **BilimALL AI Chat** - Connecting Computer Science topics with other subjects
   - Enter a topic to find related topics from okulyk.kz
   - Shows subject, grade, page number, author, publisher

2. **Slide Generator** - Presentation generation
   - Creates 5-7 slides showing CS + natural science connections
   - Downloaded in PowerPoint format

3. **Exercise Generator** - Creating tests/exercises
   - Formats: Test, Exercises, Open Questions, Game
   - Downloaded in Word format
   - Matched to grade level

Supported subjects:
- Computer Science (grades 5-11)
- Physics, Mathematics (Algebra, Geometry), Chemistry, Biology, Geography

How to use:
1. Select the desired service from the home page
2. Enter a topic (e.g. "Python loops")
3. AI will automatically generate a result

Notes:
- All content is in Kazakh
- Data is sourced only from okulyk.kz
- Results can be downloaded

Your role:
- Give short, clear answers
- Use emojis
- Explain services to users
- Provide step-by-step instructions when needed
- For off-topic questions, say "I can only provide information about BilimALL AI"

Response style: Friendly, professional, concise"""


@csrf_exempt
def chatbot_assistant(request):
    """Chatbot assistant - answers questions about BilimALL AI"""
    if request.method != "POST":
        return JsonResponse(
            {"error": "⚠️ Only POST requests are supported!"},
            status=400
        )

    try:
        data = json.loads(request.body)
        user_message = data.get("message", "").strip()

        if not user_message:
            return JsonResponse(
                {"response": "⚠️ Message cannot be empty!"},
                status=400
            )

        cache_key = f"chatbot_{hash(user_message)}"
        cached_response = cache.get(cache_key)

        if cached_response:
            return JsonResponse({"response": cached_response})

        response_text = stream_gemini_response(
            user_message,
            system_instruction=CHATBOT_SYSTEM_INSTRUCTION
        )

        cache.set(cache_key, response_text, 1800)

        return JsonResponse({"response": response_text})

    except json.JSONDecodeError:
        return JsonResponse(
            {"error": "⚠️ Invalid JSON format!"},
            status=400
        )

    except Exception as e:
        return JsonResponse(
            {"error": f"⚠️ Error: {str(e)}"},
            status=500
        )


# ============ PAGE VIEWS ============
def index(request):
    return render(request, "steam/index.html")


def chat(request):
    return render(request, "steam/chat.html")


def slide1(request):
    return render(request, "steam/slide.html")


def exercise(request):
    return render(request, "steam/exercise.html")


def fizika(request):
    return render(request, "steam/fizika.html")


def biology(request):
    return render(request, "steam/biology.html")


def geography(request):
    return render(request, "steam/geography.html")


def matem(request):
    return render(request, "steam/matem.html")


def ximia(request):
    return render(request, "steam/ximia.html")


def informatika(request):
    return render(request, "steam/informatika.html")