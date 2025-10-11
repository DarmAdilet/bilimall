import os
import json
from django.http import JsonResponse
from django.views.decorators.csrf import csrf_exempt
from google import genai
from google.genai import types
from dotenv import load_dotenv
from django.shortcuts import render
from django.conf import settings
import uuid
import re
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor

# .env файлын жүктеу
load_dotenv()
# API кілтті жүктеу
API_KEY = os.getenv("GEMINI_API_KEY")
if not API_KEY:
    raise ValueError("⚠️ GEMINI_API_KEY .env ішінде табылмады! API кілтті тексеріңіз.")

# Gemini API клиентін орнату
client = genai.Client(api_key=API_KEY)

# Модель мен генерация конфигурациясы
MODEL_NAME = "gemini-2.0-flash"
generation_config = types.GenerateContentConfig(
    temperature=0.3,
    top_p=0.95,
    top_k=40,
    max_output_tokens=8192,
    safety_settings=[
        types.SafetySetting(category="HARM_CATEGORY_HARASSMENT", threshold="BLOCK_NONE"),
        types.SafetySetting(category="HARM_CATEGORY_HATE_SPEECH", threshold="BLOCK_NONE"),
        types.SafetySetting(category="HARM_CATEGORY_SEXUALLY_EXPLICIT", threshold="BLOCK_NONE"),
        types.SafetySetting(category="HARM_CATEGORY_DANGEROUS_CONTENT", threshold="BLOCK_NONE"),
        types.SafetySetting(category="HARM_CATEGORY_CIVIC_INTEGRITY", threshold="BLOCK_NONE"),
    ],
    response_mime_type="text/plain",
    system_instruction=[
        types.Part.from_text(text="""Рөл: Сен мұғалімдерге Информатика пәнінің тақырыптарын басқа жаратылыстану пәндерімен байланыстыруға көмектесетін ассистентсің.

Нұсқаулық:

Мен саған Информатика пәнінен белгілі бір тақырыпты беремін. Сен тек https://okulyk.kz/ сайтындағы қазақша Алгебра, Геометрия, Физика, Химия, Биология, География оқулықтарын қолдана отырып, сол тақырыптың Алгебра, Геометрия, Физика, Химия, Биология, География пәндерінен (5-11 сыныптар) қандай тақырыптармен байланысы бар екенін анықта.

Байланысты тақырыптарды келесі форматта көрсет:
• Пән атауы:
• Сыныбы:
• Оқулықтың авторы және баспасы:
• Тақырыптың атауы:
• Беті:
• Байланысы:

Маңызды:
Тек нақты байланысы бар пәндерді ғана көрсет.
Бір тақырыпты бірнеше рет әртүрлі құрылғылардан берген кезде жауаптар барлық жағдайда бірдей болуы керек, артық немесе әртүрлі жауаптар берме.
Сұрақ тақырыптан ауытқыса немесе басқа салаға қатысты болса, жауап берме.
Жауап бермес бұрын оқулықтарды толық зерттеп, тек содан кейін ғана жауап қайтар.
"""),
    ],
)


@csrf_exempt
def chat_with_gemini(request):
    """Пайдаланушының хабарламасын өңдеп, Gemini-ден жауап қайтарады."""
    if request.method == "POST":
        try:
            data = json.loads(request.body)
            user_message = data.get("message", "").strip()

            if not user_message:
                return JsonResponse({"response": "⚠️ Хабарлама бос болмауы керек!"}, status=400)

            # Пайдаланушы сұранысын дайындау
            contents = [
                types.Content(role="user", parts=[types.Part.from_text(text=user_message)])
            ]

            # Жауап генерациялау
            response_text = ""
            for chunk in client.models.generate_content_stream(model=MODEL_NAME, contents=contents, config=generation_config):
                response_text += chunk.text

            return JsonResponse({"response": response_text})

        except json.JSONDecodeError:
            return JsonResponse({"error": "⚠️ JSON форматы қате!"}, status=400)

        except Exception as e:
            return JsonResponse({"error": f"⚠️ Серверлік қате: {str(e)}"}, status=500)

    return JsonResponse({"error": "⚠️ Тек POST сұранысы қолдау табады!"}, status=400)



def generate_slide_content(topic):
    contents = [
        types.Content(
            role="user",
            parts=[types.Part.from_text(
                text=(
                    f"Келесі жолда информатика және жаратылыстану пәндеріне қатысты ақпарат берілген:\n\n"
                    f"{topic}\n\n"
                    f"Бұл жолда: информатика тақырыбы, жаратылыстану пәнінің атауы, сынып, жаратылыстану пәніндегі тақырып үтір арқылы берілген.\n\n"
                    f"Сенің тапсырмаң:\n"
                    f"1. Берілген мәліметтегі информатика және жаратылыстану пәндері арасындағы байланысты түсіндір.\n"
                    f"2. Осы байланыс негізінде қазақ тілінде 5–7 слайдтан тұратын презентация мәтінін құрастыр.\n"
                    f"3. Әр слайд нақты мазмұнды болсын (мысалы: кіріспе, негізгі ұғымдар, пәндік байланыс, мысал, қорытынды).\n"
                    f"4. Ақпаратты тек okulyk.kz сайтындағы ресми оқулықтарға сүйеніп жаса.\n"
                    f"5. Әр слайдты '**N-слайд**' деген белгімен бөліп жаз.\n"
                    f"6. Әр слайдтың сөйлемдері көлемді, әрі түсінікті, нақты болсын\n\n"
                    f"7. Пайдаланылған әдебиеттер тізімін,сыныбын,баспасын, авторларын жаз\n"

                )
            )],
        ),
    ]

    generate_content_config = types.GenerateContentConfig(
        response_mime_type="text/plain",
        system_instruction=[
            types.Part.from_text(text="Тек дайын презентация мәтінін қайтар.")
        ],
    )

    response_text = ""
    for chunk in client.models.generate_content_stream(model=MODEL_NAME, contents=contents, config=generate_content_config):
        response_text += chunk.text

    return response_text.strip()

def create_kazakh_slides(prs, topic, content):
    """Қазақ тіліндегі бірнеше слайдтарды құру"""
    # Слайдтарды бөлу
    slides_content = re.split(r'\*\*(\d+-слайд.*?)\*\*', content)

    # Бірінші слайд (тақырып слайды)
    title_slide_layout = prs.slide_layouts[0]
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.title
    subtitle = slide.placeholders[1]
    title.text = topic
    subtitle.text = "Автоматты түрде жасалған презентация"

    # Қалған слайдтар
    for i in range(1, len(slides_content), 2):
        if i + 1 >= len(slides_content):
            continue

        slide_title = slides_content[i].strip()
        slide_content = slides_content[i + 1].strip()

        if not slide_content:
            continue

        # Жаңа слайд құру
        content_slide_layout = prs.slide_layouts[1]
        slide = prs.slides.add_slide(content_slide_layout)

        # Тақырыпты пішімдеу
        title_shape = slide.shapes.title
        title_shape.text = slide_title
        title_shape.text_frame.paragraphs[0].font.bold = True
        title_shape.text_frame.paragraphs[0].font.size = Pt(28)
        title_shape.text_frame.paragraphs[0].font.color.rgb = RGBColor(0, 84, 149)
        title_shape.text_frame.paragraphs[0].font.name = 'Arial'

        # Мазмұнды пішімдеу
        body_shape = slide.placeholders[1]
        tf = body_shape.text_frame
        tf.clear()

        # Параграфтарға бөлу
        for paragraph in slide_content.split('\n'):
            if paragraph.strip():
                p = tf.add_paragraph()
                p.text = paragraph.strip()
                p.font.size = Pt(18)
                p.font.name = 'Arial'
                p.space_after = Pt(8)

                # Тармақтарды белгілеу
                if any(paragraph.strip().startswith(char) for char in ['•', '-', '1.', '2.', '3.']):
                    p.level = 1
                    p.font.size = Pt(16)


@csrf_exempt
def generate_slide(request):
    """Слайдты генерациялау API"""
    if request.method != 'POST':
        return JsonResponse({'error': 'Тек POST сұраныстар қабылданады'}, status=405)

    try:
        if request.content_type == 'application/json':
            data = json.loads(request.body)
            topic = data.get('topic', '').strip()
        else:
            topic = request.POST.get('topic', '').strip()

        if not topic:
            return JsonResponse({'error': 'Тақырып бос болмауы керек'}, status=400)

        # Контент генерациялау
        slide_text = generate_slide_content(topic)

        if not slide_text:
            return JsonResponse({'error': 'Контент генерацияланбады'}, status=500)

        # Презентация жасау
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

    except json.JSONDecodeError:
        return JsonResponse({'error': 'Жарамсыз JSON'}, status=400)

    except Exception as e:
        return JsonResponse({
            'error': f'Қате орын алды: {str(e)}',
            'slide_text': 'Слайд жасау кезінде қате болды.'
        }, status=500)
def index(request):
    """Чаттың UI бетін көрсету."""
    return render(request, "steam/index.html")

def chat(request):
    """Чаттың UI бетін көрсету."""
    return render(request, "steam/chat.html")

def slide1(request):
    return render(request,"steam/slide.html")

def fizika(request):
    return render(request,"steam/fizika.html")

def biology(request):
    return render(request,"steam/biology.html")

def geography(request):
    return render(request,"steam/geography.html")

def matem(request):
    return render(request,"steam/matem.html")

def ximia(request):
    return render(request,"steam/ximia.html")

def informatika(request):
    return render(request,"steam/informatika.html")




