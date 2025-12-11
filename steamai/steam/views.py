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

# PowerPoint үшін
from pptx import Presentation
from pptx.util import Pt as PptxPt
from pptx.dml.color import RGBColor as PptxRGBColor

# Word үшін
from docx import Document
from docx.shared import Pt as DocxPt, RGBColor as DocxRGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ============ КОНФИГУРАЦИЯ ============
load_dotenv()
API_KEY = os.getenv("GEMINI_API_KEY")

if not API_KEY:
    raise ValueError("⚠️ GEMINI_API_KEY .env ішінде табылмады!")

client = genai.Client(api_key=API_KEY)
MODEL_NAME = "gemini-2.5-flash"

# Қауіпсіздік параметрлері
SAFETY_SETTINGS = [
    types.SafetySetting(category="HARM_CATEGORY_HARASSMENT", threshold="BLOCK_NONE"),
    types.SafetySetting(category="HARM_CATEGORY_HATE_SPEECH", threshold="BLOCK_NONE"),
    types.SafetySetting(category="HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold="BLOCK_NONE"),
    types.SafetySetting(category="HARM_CATEGORY_DANGEROUS_CONTENT", threshold="BLOCK_NONE"),
    types.SafetySetting(category="HARM_CATEGORY_CIVIC_INTEGRITY", threshold="BLOCK_NONE"),
]

# Чат үшін арнайы system instruction
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
Жауап бермес бұрын оқулықтарды толық зерттеп, тек содан кейін ғана жауап қайтар."""

# Негізгі генерация конфигурациясы (чат үшін)
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


# ============ HELPER ФУНКЦИЯЛАР ============
def get_generation_config(system_instruction=None):
    """Генерация конфигурациясын құру"""
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
    """Gemini API-дан жауап алу"""
    # Егер чат конфигурациясын қолдану керек болса
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
        raise Exception(f"Gemini API қатесі: {str(e)}")

    return response_text.strip()


def validate_json_request(request):
    """JSON сұранысын тексеру"""
    if request.method != 'POST':
        return None, JsonResponse(
            {'error': 'Тек POST сұраныстар қабылданады'},
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
            {'error': 'Жарамсыз JSON форматы'},
            status=400
        )


# ============ ЧАТ ============
@csrf_exempt
def chat_with_gemini(request):
    """Пайдаланушының хабарламасын өңдеп, Gemini-ден жауап қайтарады."""
    if request.method != "POST":
        return JsonResponse(
            {"error": "⚠️ Тек POST сұранысы қолдау табады!"},
            status=400
        )

    try:
        data = json.loads(request.body)
        user_message = data.get("message", "").strip()

        if not user_message:
            return JsonResponse(
                {"response": "⚠️ Хабарлама бос болмауы керек!"},
                status=400
            )

        # Кэштен тексеру
        cache_key = f"chat_{hash(user_message)}"
        cached_response = cache.get(cache_key)

        if cached_response:
            return JsonResponse({"response": cached_response})

        # Жауап генерациялау (чат конфигурациясын қолдану)
        response_text = stream_gemini_response(user_message, use_chat_config=True)

        # Кэшке сақтау
        cache.set(cache_key, response_text, 3600)

        return JsonResponse({"response": response_text})

    except json.JSONDecodeError:
        return JsonResponse(
            {"error": "⚠️ JSON форматы қате!"},
            status=400
        )

    except Exception as e:
        return JsonResponse(
            {"error": f"⚠️ Серверлік қате: {str(e)}"},
            status=500
        )


# ============ WORD ЖҮКТЕУ ============
@csrf_exempt
def download_exercise_word(request):
    """Тапсырманы Word файлына жүктеу"""
    data, error_response = validate_json_request(request)
    if error_response:
        return error_response

    exercise_content = data.get('content', '').strip()
    topic = data.get('topic', 'Тапсырма').strip()
    level = data.get('level', '').strip()
    format_type = data.get('format', '').strip()

    if not exercise_content:
        return JsonResponse(
            {'error': 'Тапсырма мазмұны жоқ'},
            status=400
        )

    try:
        # Word документін құру
        doc = Document()

        # Стильдер
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Times New Roman'
        font.size = DocxPt(12)

        # Тақырып
        title = doc.add_heading(topic, level=0)
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title.runs[0]
        title_run.font.size = DocxPt(18)
        title_run.font.bold = True
        title_run.font.color.rgb = DocxRGBColor(0, 84, 149)

        # Метадеректер
        doc.add_paragraph()
        meta_para = doc.add_paragraph()
        meta_para.add_run(f'📊 Сынып: ').bold = True
        meta_para.add_run(level)
        meta_para.add_run('\n📝 Формат: ').bold = True
        meta_para.add_run(format_type)
        meta_para.add_run('\n📅 Күні: ').bold = True
        meta_para.add_run(datetime.now().strftime('%Y-%m-%d'))

        # Сызық
        doc.add_paragraph('_' * 60)
        doc.add_paragraph()

        # Мазмұн
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

        # Footer
        doc.add_paragraph()
        footer_para = doc.add_paragraph()
        footer_run = footer_para.add_run('\n\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n')
        footer_run.font.color.rgb = DocxRGBColor(128, 128, 128)
        footer_para.add_run('Генерацияланған: BilimAll AI\n').italic = True
        footer_para.add_run('https://bilimall.kz').font.color.rgb = DocxRGBColor(0, 84, 149)
        footer_para.alignment = WD_ALIGN_PARAGRAPH.CENTER

        # Файлды жадта сақтау
        doc_io = io.BytesIO()
        doc.save(doc_io)
        doc_io.seek(0)

        # HTTP жауабы
        response = HttpResponse(
            doc_io.read(),
            content_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

        filename = f"{topic.replace(' ', '_')}_тапсырма.docx"
        response['Content-Disposition'] = f'attachment; filename="{filename}"'

        return response

    except Exception as e:
        print(f"Word қатесі: {str(e)}")
        return JsonResponse({
            'error': f'Word файлын жасау кезінде қате: {str(e)}'
        }, status=500)


# ============ ОЙЫН ФУНКЦИЯЛАРЫ ============
def generate_game_questions(topic, level):
    """Ойын сұрақтарын генерациялау"""
    prompt = f"""Python бағдарламалау тілін үйренуге арналған интерактивті викторина сұрақтарын құрастыр.

Параметрлер:
- Тақырып: {topic}
- Деңгей: {level}

Нұсқаулық:
1. 10 сұрақ жаса
2. Әр сұрақта 4 жауап нұсқасы болсын
3. Сұрақтар Python синтаксисі, функциялар, циклдер, шарттар, деректер түрлері туралы болсын
4. Кейбір сұрақтарда код мысалдары көрсет
5. Дұрыс жауап нөмірін көрсет (0-3 аралығында)
6. Әр жауапқа қысқа түсініктеме бер

МІНДЕТТІ: Жауабыңызды тек JSON форматында қайтар.

JSON форматы:
{{
  "questions": [
    {{
      "question": "Сұрақ мәтіні",
      "code": "print('Hello')",
      "options": ["A нұсқа", "B нұсқа", "C нұсқа", "D нұсқа"],
      "correctIndex": 0,
      "explanation": "Түсініктеме"
    }}
  ]
}}

Ескертпе: "code" өрісі опционалды."""

    system_inst = "Сен Python үйретуші ботсың. Тек JSON форматында жауап қайтар!"

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
        print(f"JSON parse қатесі: {e}")
        return []


@csrf_exempt
def generate_game_questions_api(request):
    """Ойын сұрақтарын генерациялау API"""
    data, error_response = validate_json_request(request)
    if error_response:
        return error_response

    topic = data.get('topic', '').strip()
    level = data.get('level', '').strip()

    if not topic or not level:
        return JsonResponse(
            {'error': 'Тақырып және деңгей көрсетілуі керек'},
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
                {'error': 'Сұрақтар генерацияланбады'},
                status=500
            )

        cache.set(cache_key, questions, 1800)

        return JsonResponse({
            'success': True,
            'questions': questions
        })

    except Exception as e:
        return JsonResponse({
            'error': f'Қате орын алды: {str(e)}'
        }, status=500)


def game_view(request):
    """Ойын бетін көрсету"""
    topic = request.GET.get('topic', 'Python негіздері')
    level = request.GET.get('level', '5-сынып')

    return render(request, 'steam/game.html', {
        'topic': topic,
        'level': level
    })


# ============ ТАПСЫРМА ГЕНЕРАЦИЯЛАУ ============
def generate_exercise_content(topic, format_type, level):
    """Тапсырма мазмұнын генерациялау"""
    prompt = f"""Келесі параметрлерге сәйкес тапсырма құрастыр:

Тақырып: {topic}
Формат: {format_type}
Сынып деңгейі: {level}

Нұсқаулық:
1. Тапсырма қазақ тілінде болуы керек
2. Сынып деңгейіне сәйкес қиындық болсын
3. Таңдалған форматқа сәйкес жаса:
   - Тест: 20 сұрақ, 4 нұсқалы жауап, дұрыс жауаптар көрсетілсін
   - Тапсырма: 5-7 практикалық тапсырма, нұсқаулықтармен
   - Сұрақтар: 8-10 ашық сұрақ, әртүрлі қиындық деңгейімен

4. Тапсырмалар нақты, түсінікті және білім бағалауға бағытталған болсын"""

    system_inst = "Сен тәжірибелі мұғалімсің. Тек дайын тапсырманы қайтар."

    return stream_gemini_response(prompt, system_inst)


@csrf_exempt
def generate_exercise(request):
    """Тапсырма генерациялау API"""
    data, error_response = validate_json_request(request)
    if error_response:
        return error_response

    topic = data.get('topic', '').strip()
    format_type = data.get('format', '').strip()
    level = data.get('level', '').strip()

    if not topic or not format_type or not level:
        return JsonResponse(
            {'error': 'Барлық өрістер толтырылуы керек'},
            status=400
        )

    if format_type == 'Ойын':
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
                {'error': 'Тапсырма генерацияланбады'},
                status=500
            )

        cache.set(cache_key, exercise_content, 1800)

        return JsonResponse({
            'success': True,
            'exercise_content': exercise_content
        })

    except Exception as e:
        return JsonResponse({
            'error': f'Қате орын алды: {str(e)}'
        }, status=500)


# ============ СЛАЙД ГЕНЕРАЦИЯЛАУ ============
def create_slide_text_frame(shape, content, title_level=False):
    """Слайд мәтінін пішімдеу"""
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
    """Слайд мазмұнын генерациялау"""
    prompt = f"""Келесі жолда информатика және жаратылыстану пәндеріне қатысты ақпарат берілген:

{topic}

Бұл жолда: информатика тақырыбы, жаратылыстану пәнінің атауы, сынып, жаратылыстану пәніндегі тақырып үтір арқылы берілген.

Сенің тапсырмаң:
1. Берілген мәліметтегі информатика және жаратылыстану пәндері арасындағы байланысты түсіндір.
2. Осы байланыс негізінде қазақ тілінде 5–7 слайдтан тұратын презентация мәтінін құрастыр.
3. Әр слайд нақты мазмұнды болсын (мысалы: кіріспе, негізгі ұғымдар, пәндік байланыс, мысал, қорытынды).
4. Ақпаратты тек okulyk.kz сайтындағы ресми оқулықтарға сүйеніп жаса.
5. Әр слайдты '**N-слайд**' деген белгімен бөліп жаз.
6. Әр слайдтың сөйлемдері көлемді, әрі түсінікті, нақты болсын
7. Пайдаланылған әдебиеттер тізімін,сыныбын,баспасын, авторларын жаз"""

    system_inst = "Тек дайын презентация мәтінін қайтар."
    return stream_gemini_response(prompt, system_inst)


def create_kazakh_slides(prs, topic, content):
    """Қазақ тіліндегі слайдтарды құру"""
    slides_content = re.split(r'\*\*(\d+-слайд.*?)\*\*', content)

    # Бірінші слайд (тақырып)
    title_slide = prs.slides.add_slide(prs.slide_layouts[0])
    title_slide.shapes.title.text = topic
    title_slide.placeholders[1].text = "Автоматты түрде жасалған презентация"

    # Қалған слайдтар
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
    """Слайд генерациялау API"""
    data, error_response = validate_json_request(request)
    if error_response:
        return error_response

    topic = data.get('topic', '').strip()

    if not topic:
        return JsonResponse(
            {'error': 'Тақырып бос болмауы керек'},
            status=400
        )

    try:
        slide_text = generate_slide_content(topic)

        if not slide_text:
            return JsonResponse(
                {'error': 'Контент генерацияланбады'},
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
            'error': f'Қате орын алды: {str(e)}',
            'slide_text': 'Слайд жасау кезінде қате болды.'
        }, status=500)
# Ассистент помощник
CHATBOT_SYSTEM_INSTRUCTION = """Сен BilimALL AI жобасының көмекші ассистентісің.

BilimALL AI туралы ақпарат:
- Бұл – информатика мен жаратылыстану пәндерін байланыстыратын AI платформасы
- Қ.Жұбанов атындағы Ақтөбе өңірлік университетінде жасалған
- 5-11 сыныптарға арналған

Қызметтер:
1. **BilimALL AI Chat** - Информатика тақырыптарын басқа пәндермен байланыстыру
   - Тақырыпты енгізсе, okulyk.kz сайтынан байланысты тақырыптарды табады
   - Пән, сынып, бет нөмірі, автор, баспа көрсетеді

2. **Slide Generator** - Презентация генерациясы
   - Информатика + жаратылыстану пәні байланысын көрсететін 5-7 слайд жасайды
   - PowerPoint форматында жүктеледі

3. **Тапсырма генераторы** - Тест/тапсырмалар құрастыру
   - Форматтар: Тест, Тапсырмалар, Ашық сұрақтар, Ойын
   - Word форматында жүктеледі
   - Сынып деңгейіне сәйкес

Қолдау көрсетілетін пәндер:
- Информатика (5-11 сынып)
- Физика, Математика (Алгебра, Геометрия), Химия, Биология, География

Пайдалану нұсқаулығы:
1. Басты беттен қажетті қызметті таңда
2. Тақырып енгіз (мысалы: "Python циклдері")
3. AI автоматты түрде нәтиже береді

Ескертпелер:
- Барлық контент қазақ тілінде
- Тек okulyk.kz сайтынан дерек алады
- Нәтижелерді жүктеп алуға болады

Сенің міндетің:
- Қысқа, нақты жауап бер
- Эмодзи қолдан
- Пайдаланушыға қызметтерді түсіндір
- Қажет болса қадамдық нұсқаулық бер
- Тақырыптан тыс сұрақтарға "Мен тек BilimALL AI туралы ақпарат бере аламын" де

Жауап стилі: Достық, кәсіби, қысқа"""

@csrf_exempt
def chatbot_assistant(request):
    """Чатбот ассистент - BilimALL AI туралы сұрақтарға жауап береді"""
    if request.method != "POST":
        return JsonResponse(
            {"error": "⚠️ Тек POST сұранысы қолдау табады!"},
            status=400
        )

    try:
        data = json.loads(request.body)
        user_message = data.get("message", "").strip()

        if not user_message:
            return JsonResponse(
                {"response": "⚠️ Хабарлама бос болмауы керек!"},
                status=400
            )

        # Кэштен тексеру
        cache_key = f"chatbot_{hash(user_message)}"
        cached_response = cache.get(cache_key)

        if cached_response:
            return JsonResponse({"response": cached_response})

        # Gemini-ден жауап алу
        response_text = stream_gemini_response(
            user_message,
            system_instruction=CHATBOT_SYSTEM_INSTRUCTION
        )

        # Кэшке сақтау (30 минут)
        cache.set(cache_key, response_text, 1800)

        return JsonResponse({"response": response_text})

    except json.JSONDecodeError:
        return JsonResponse(
            {"error": "⚠️ JSON форматы қате!"},
            status=400
        )

    except Exception as e:
        return JsonResponse(
            {"error": f"⚠️ Қате: {str(e)}"},
            status=500
        )
# ============ БЕТТІ КӨРСЕТУ ============
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