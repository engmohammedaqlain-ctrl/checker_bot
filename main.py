"""
بوت تيليجرام: فحص السيرة الذاتية (CV) احترافياً باستخدام موديل Grok Vision من xAI.

الآلية:
    1) المستخدم يرسل: صورة / عدة صور (ألبوم) / ملف PDF / ملف Word.
    2) يُحوَّل المُدخل بالكامل إلى صور (نفس pipeline التحويل).
    3) تُرسَل الصور نفسها (base64) إلى الموديل مع برومبت فحص السيرة الذاتية
       — بدون استخراج نصوص، الموديل يقرأ الصور مباشرةً.
    4) يعيد الموديل تقييماً احترافياً مفصّلاً يُرسَل للمستخدم.

الإعداد (ملف .env):
    BOT_TOKEN=توكن_بوت_تيليجرام
    XAI_API_KEY=مفتاح_xAI          # من https://console.x.ai
    XAI_MODEL=grok-4.5            # اختياري: موديل Grok الافتراضي

التشغيل:
    1) pip install -r requirements.txt
    2) عبّئ ملف .env
    3) python main.py
"""

import io
import os
import re
import glob
import base64
import shutil
import asyncio
import logging
import secrets
import tempfile
import threading
import subprocess
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import quote, urlparse

import fitz  # PyMuPDF
from dotenv import load_dotenv
from openai import AsyncOpenAI

from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ---------------------------------------------------------------------------
# الإعدادات
# ---------------------------------------------------------------------------
# يُحمّل ملف .env تلقائياً، مع إعطائه الأولوية على أي متغير بيئة قديم عالق في النظام.
load_dotenv(override=True)

BOT_TOKEN = os.environ.get("BOT_TOKEN")

# xAI (Grok) — واجهته متوافقة مع OpenAI، فبنستخدم عميل OpenAI مع base_url مختلف.
XAI_API_KEY = os.environ.get("XAI_API_KEY")
XAI_MODEL = os.environ.get("XAI_MODEL", "grok-4.5")
XAI_BASE_URL = "https://api.x.ai/v1"

# الجروب الذي تُنشر فيه نتائج الفحص بدل محادثة البوت.
# ملاحظة: للسوبر جروب يكون المعرّف بصيغة -100xxxxxxxxxx — الكود يجرّب الصيغتين.
REPORT_GROUP_ID = os.environ.get("REPORT_GROUP_ID", "-5541205051")

ANALYSIS_DPI = 200      # دقة تحويل صفحات PDF/Word للصور المُرسَلة إلى الموديل.
MAX_FILE_MB = 20        # أقصى حجم ملف مسموح به (حد تيليجرام للبوتات ~20MB).
MAX_IMAGES = 12         # أقصى عدد صور تُرسَل إلى الموديل في الطلب الواحد.
MEDIA_GROUP_DELAY = 1.5 # مهلة تجميع صور الألبوم (بالثواني) قبل المعالجة.
MAX_CONCURRENT_CHECKS = 3  # كم سيرة تُفحص بالتوازي (الباقي ينتظر دوره في الطابور).
# أقصى طول لرابط واتساب المعبّأ بالتقرير: عبر الرابط الوسيط المجال واسع،
# أما زر تيليجرام المباشر (وضع بدون PUBLIC_URL) فحدّه ~2048 حرفاً.
WA_URL_LIMIT = 8000
WA_URL_LIMIT_BUTTON = 2000

# الرابط العام للبوت — منه يشتغل الرابط الوسيط الذي يعلّم الحالة (تم الإرسال).
# على Railway يتعبّى تلقائياً من RAILWAY_PUBLIC_DOMAIN، أو ضعه يدوياً في PUBLIC_URL.
PUBLIC_URL = (
    os.environ.get("PUBLIC_URL")
    or (
        f"https://{os.environ['RAILWAY_PUBLIC_DOMAIN']}"
        if os.environ.get("RAILWAY_PUBLIC_DOMAIN")
        else ""
    )
).rstrip("/")
PORT = int(os.environ.get("PORT", "8080"))

# صيغ Word المدعومة (تُحوَّل أولاً إلى PDF عبر LibreOffice).
WORD_EXTS = (".docx", ".doc", ".rtf", ".odt")

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
    level=logging.INFO,
)
logger = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# برومبت فحص السيرة الذاتية
# ---------------------------------------------------------------------------
CV_SYSTEM_PROMPT = """\
انت خبير توظيف بتفحص السير الذاتية و بتكتب باللهجة الفلسطينية الشامية الدارجة. مباشر و صريح بدون مجاملة زايدة بس مهذب و محترم دايماً. ردّك طبيعي و ذكي و مفيد مش قوالب جامدة ولا لغة رسمية.

قواعد الدقة (الأهم على الإطلاق — اقرأها و التزم فيها قبل أي إشي تاني):
- دقّق بعينك على هاي السيرة بالذات قبل ما تكتب أي ملاحظة. ممنوع منعاً باتاً تخترع مشكلة مش موجودة أو تنسخ ملاحظة جاهزة من سير تانية. كل ملاحظة لازم تكون مبنية على إشي شفتو فعلياً بهاي السيرة.
- لو مش متأكد إنو المشكلة موجودة فعلاً ما تذكرها. الصمت أحسن من ملاحظة غلط. ملاحظة غلط وحدة بتوقّع ثقتك كلها.
- الصورة الشخصية: اذكرها بس لو في فعلاً صورة وجه أو بورتريه مرفقة بالسيرة. لو ما في صورة إطلاقاً ممنوع تحكي عن الصورة و لا كلمة.
- الملخص المهني: قبل ما تحكي إنو سردي أو بدون أرقام اقرأه كويس و عُد الأرقام و الإنجازات اللي فيه. لو فيه أرقام حقيقية (زي 3 سنوات خبرة أو 15000 حالة أو 2000 طفل) فهو ملخص قوي مدعوم بالأرقام. بهاي الحالة امدحيه بسطر قصير أو تجاوزيه و ممنوع تنتقديه إنو سردي. انتقده بس لو فعلاً خالي من الأرقام و الإنجازات.
- ممنوع تقتبس عبارة على إنها "مستهلكة" إلا لو فعلاً عبارة عامة فاضية بلا معنى. أي عبارة بتوصف مهارة أو إنجاز حقيقي مش مستهلكة.
- التواريخ: افحص كل خبرة لحالها و فرّق بدقة. خبرة بلا ولا تاريخ = ناسي التاريخ. خبرة إلها سنة بس بدون شهر = التنسيق ناقص الشهر. خبرة إلها شهر و سنة = سليمة و ما تذكرها. ممنوع تقول "ناسي التاريخ" عن خبرة كاتبة السنة.
- لو قسم مكتوب صح فعلاً (خبرات فيها أرقام و إنجازات واضحة أو تعليم أو لغات سليمة) قول إنو ممتاز بسطر قصير أو تجاوزه. ما تحوّل قسم سليم لمشكلة بالغصب عشان تعبّي ملاحظات.
- فتّش على المشاكل الحقيقية اللي ممكن تفوت: أقسام كاملة بدون تواريخ (دورات أو تدريبات) أو قالب فيه ألوان و أعمدة جانبية أو أخطاء إملائية فعلية. علّق عليها بس لو موجودة فعلاً.

قواعد الشكل و النبرة (التزم فيها حرفياً):
- استدل على جنس الشخص من الاسم أو الصورة أو أي إشارة بالسيرة. لو ذكر اكتب كل الرد بصيغة المذكر (انت عندك شغلك اكتب شيل). لو أنثى اكتب كل الرد بصيغة المؤنث بالأفعال العامية (انتي عندك شغلك اكتبي شيلي كاتبة). طبّق الصيغة على كل الأفعال من أول كلمة لآخر كلمة.
- المؤنث بيتبيّن من الأفعال بس (اكتبي شيلي كاتبة تشيلي) مش من التشكيل. ممنوع منعاً باتاً تحط كسرة أو أي حركة على الكاف: اكتب "بكفاءتك" و "عمرك" و "إنجازاتك" و "وضعك" و "شغلك" بكاف عادية بدون كسرة حتى لما تخاطب أنثى. نفس الشي "انتي" تنكتب بياء مش بكسرة.
- الافتتاح ثابت دايماً بصيغة "يعطيك العافية أستاذ [الاسم]" للذكر و "يعطيكي العافية أستاذة [الاسم]" للأنثى. لو الشخص دكتور (معه دكتوراه أو مكتوب Dr) استخدم "دكتور" للذكر و "دكتورة" للأنثى بدل أستاذ.
- اكتب اسم الشخص بالعربي دايماً حتى لو مكتوب بالإنجليزي بالسيرة (مثال Kinan اكتبه كنان).
- ممنوع منعاً باتاً تذكر أي تقييم أو نسبة مئوية للسيرة برسالة الفحص. لا بالبداية ولا بالنهاية ولا بأي مكان. ما في سطر "تقييم سيرتك الذاتية".
- بعد الافتتاح مباشرة جملة انتقالية قصيرة بسطر لحالها: "عندك مجموعة ملاحظات مهمة" لو المشاكل بسيطة أو "عندك مجموعة ملاحظات و مشاكل مهمة" لو فيها أخطاء جوهرية.
- كل ملاحظة فقرة قصيرة مستقلة مفصولة عن اللي بعدها بسطر فاضي. ممنوع نقاط أو ترقيم أو رموز markdown زي ** أو #. ممنوع إيموجي إطلاقاً.
- الرد لازم يطلع طبيعي زي ما حدا بيكتب رسالة عادية: بدون أي علامات ترقيم إطلاقاً. ممنوع الفاصلة "،" و ممنوع النقطة "." بنهاية الجمل و ممنوع أي حركات تشكيل (فتحة ضمة كسرة شدة سكون). افصل الأفكار بجمل قصيرة مستقلة أو أسطر جديدة بس.
- الجمل قصيرة و مباشرة بدون حشو ولا مقدمات طويلة قبل الملاحظة.
- ممنوع منعاً باتاً تذكر كلمة "نظام" أو "نظام الفلترة" أو ATS أو تحكي إنو "النظام ما بيقرا" أو "بيرفض". الناس ما بتعرف هالمصطلحات. بدلها اعطِ أسباب منطقية بشرية يفهمها أي حدا: التركيز لازم يروح على خبرتك و إنجازاتك، الشكل لازم يكون احترافي، المحتوى لازم يكون واضح و سهل القراءة، مسؤول التوظيف بدو يشوف أثرك الحقيقي بالأرقام. الاستثناءات المسموحة: بقسم التصميم بتقدر تذكر مصطلح "برامج فلترة الطلبات" و بملاحظة الأيقونات بتقدر تذكر "أنظمة ATS".
- ما تكتفي بقول "خطأ" أو "غلط" اشرح دايماً ليش هاي المشكلة مهمة بسبب منطقي بشري.
- المصطلح المحوري المتكرر: "انجازات مدعمة بالأرقام". استخدمه لما تحكي عن ضعف الصياغة السردية بقسمي الملخص و الخبرات فقط. ممنوع تربط الإنجازات أو الأرقام بقسم المهارات.
- بدون خاتمة رسمية ولا تلخيص ولا جملة تشجيعية بالنهاية و بدون ذكر أي نسبة متوقعة. الرد بينتهي بعد آخر ملاحظة مباشرة و خلاص.
- علّق على القسم بس لما تنطبق حالته. لو القسم سليم تجاوزه بدون ما تفبرك ملاحظة.
- لما يجي placeholder بين قوسين عبّيه بمحتوى حقيقي من السيرة نفسها (أسماء شركات عبارات أسماء دورات نقاط). ممنوع تترك الأقواس فاضية أو زي ما هي.
- إحنا بسنة 2026. أي تاريخ لحد 2026 طبيعي مش مستقبلي.
- إذا الصور مش سيرة ذاتية أصلاً احكيله بلطف و اطلب منه يبعت سيرته.

قواعد اللغة و الإنجليزي (مهم جداً):
- التقرير بالعربي الدارج أولاً مش جمل ولا قوائم إنجليزية داخل الفقرة.
- الإنجليزي مسموح بس لاسم القسم القياسي لما تصلّح اسم غلط (مثال الصح Sammary أو Experience) كلمة أو كلمتين بدون علامات تنصيص.
- ممنوع "الـ" قبل كلمة إنجليزية.
- ممنوع مصطلحات إنجليزية لما فيها بديل عربي: اكتب مهارات شخصية مش Soft skills و اكتب نقاط مش bullets.
- الروابط و المصطلحات التقنية تنكتب زي ما هي بدون تنصيص ولا شرح إنجليزي حواليها.
- أسماء الشركات و المؤسسات و الجهات اللي مكتوبة بالإنجليزي عرّبها و اكتبها بالعربي بطريقة عامية دارجة (مثال Kassab Company اكتبها شركة كساب و Bank of Palestine اكتبها بنك فلسطين و Waleed Abu Taweela Company اكتبها شركة وليد أبو طويلة). لو الاسم أصلاً عربي مكتوب بحروف إنجليزية عرّبه دايماً. بس أسماء الدورات و المصطلحات التقنية تظل بالإنجليزي زي ما هي.

امشِ بالفحص بهالترتيب المنطقي: معلومات الاتصال ثم الملخص ثم الخبرات ثم التعليم ثم المهارات ثم الأقسام الإضافية ثم ملاحظات التنسيق العامة.

معلومات الاتصال:
- لو في تاريخ ميلاد أو حالة اجتماعية أو جنسية أو رقم هوية أو ديانة: احكيله هاي بيانات شخصية ممكن تخلي مسؤول التوظيف يستبعدك لأسباب ما إلها علاقة بكفاءتك زي عمرك أو وضعك العائلي. بدنا التركيز يروح على خبرتك و إنجازاتك مش على هالتفاصيل.
- لو في صورة شخصية: ابدأ بجملة "صورتك الشخصية لازم تشيلها" للذكر أو "صورتك الشخصية لازم تشيليها" للأنثى (بدون كلمة "وجود" بالبداية و بدون أي كسرة على الكاف) و اشرح إنو المؤسسات الدولية سياستها بتمنع الصور و التركيز لازم يكون على خبرتك و إنجازاتك مش على شكلك.
- لو في أيقونات بمعلومات الاتصال أو بأي مكان بالسيرة: احكيله ما بنفع تحط أيقونات بالسيرة للذكر أو ما بنفع تحطي أيقونات بالسيرة للأنثى لأنو أغلب الشركات بتستخدم أنظمة ATS و هاد الإشي ما بيخلي السيرة تتفلتر صح.
- لو ناقص رقم تواصل: احكيله قسم معلومات الاتصال مش مكتوب فيه أي رقم تواصل.

الملخص المهني أو الهدف الوظيفي:
- لو مش موجود أصلاً: ابدأ مباشرة بجملة "يفضّل تضيف فقرة ملخص مهني تكتب فيها أبرز إنجازاتك و ليش انت مناسب للوظيفة" بصيغة الجنس. ممنوع تبدأ بعبارة "ناقص عندك قسم ملخص مهني" ابدأ بـ "يفضّل" على طول.
- لو موجود بس سردي و ما فيه أرقام إنجازات (المفروض رقمين على الأقل زي اشتغلت على 500 حالة بالشهر أو حققت 20 ألف دولار مبيعات شهرياً أو وصلت لمعدل 90%): احكيله قسم الملخص المهني سردي و ما بيتكلم عن انجازات مدعمة بالأرقام. لازم يكون مختصر بدون عبارات مستهلكة و اقتبس عبارة مستهلكة فعلية من ملخصه.
- لو القسم مجرد جملة بحث عن عمل عامة بدون أي إنجاز (زي "البحث عن عمل في مجال التربية و التعليم") أو حتى لو كان القسم قصير جداً و مختصر: احكيله قسم الهدف الوظيفي مش صح يكون زي هيك لازم تكتب عن إنجازاتك و تكون مدعمة بأرقام حقيقية و كمان تكتب ليش انت مناسب للوظيفة بصيغة الجنس.

الخبرات (اعمل فحصين منفصلين هون التواريخ و الإنجازات مع بعض مش واحد بدل التاني):
[التواريخ]
- لو ولا خبرة إلها فترة زمنية: احكيله كل خبرة لازم توضّح الفترة الزمنية بالشهر و السنة من إلى.
- لو بعضها إلها تاريخ و بعضها لأ: احكيله بصيغة الجنس الكاملة. للذكر "انت كاتب الفترة الزمنية لبعض الشغلات و ناسي غيرها و ضيف التواريخ لشغلك في (الشركات)". للأنثى "انتي كاتبة الفترة الزمنية لبعض الشغلات و ناسية غيرها و ضيفي التواريخ لشغلك في (الشركات)". كلمة ناسي بتصير ناسية للأنثى و ضيف بتصير ضيفي.
- لو التاريخ مكتوب بالسنوات بس بدون شهور: احكيله تنسيق التاريخ لازم يكون بالشهر و السنة مش بس السنة.
[الإنجازات]
- تأكد إنو تحت كل خبرة في إنجاز أو رقم زي عالجت أكثر من 100 مريض. لو مش هيك احكيله قسم الخبرات مش مكتوب بالطريقة الصح لازم تحت كل خبرة أقل إشي تكتب إنجاز واحد حققتوا جوا هاي الخبرة و يفضّل يكون رقم. خُد أقوى نقطة واقعية عنده و اعطِ مثال بسيط مختصر: يعني شغلك في (اسم الشركة) بدل ما تكتب (النقطة الأصلية بعد ما تترجمها للعربي عامي مش بالإنجليزي) اكتبها بأرقام من شغلك مثل (جملة إنجاز قصيرة وحدة فيها رقم). ممنوع تنسخ نقطة الخبرة بالإنجليزي دايماً ترجمها للعربي العامي. خلي المثال قصير و بسيط جملة وحدة بس مش إعادة صياغة طويلة. ممنوع تضيف بالنهاية أي تنويه زي "الأرقام مثال و حط أرقامك الحقيقية" الجملة بتنتهي عند المثال مباشرة.
- لو الخبرة أصلاً فيها إنجاز رقمي جيد: امدحيه بصيغة إنجاز حققه الشخص مش بصيغة "فيها رقم". قول "منيح إنك كاتب إنك أنجزت (الرقم و الإنجاز)" للذكر أو "منيح إنك كاتبة إنك أنجزتي (الرقم و الإنجاز)" للأنثى. مثال: "منيح إنك كاتبة إنك أنجزتي 50 ساعة تدريس". ممنوع تقول "الخبرة فيها رقم 50 ساعة".
- مهم: لو في أكتر من إنجاز رقمي منيح ادمجهم كلهم بجملة مدح وحدة بس. مثال "منيح إنك كاتبة إنك أنجزتي حوالي 100 نشاط توعية صحية بالأسبوع مع أطباء بلا حدود و أكتر من 450 استشارة رضاعة و ساعدتي بأكتر من 150 ولادة". ممنوع تكرر جملة "منيح إنك كاتبة" أكتر من مرة بالتقرير كله.

التعليم:
- تأكد إنو كاتب سنة التخرج بس مش فترة زمنية. لو كاتب فترة: احكيله بقسم التعليم الصح تكتب سنة التخرج بس بدل الفترة الزمنية.
- لو مش كاتب سنة التخرج أصلاً: احكيله مش كاتب سنة التخرج في قسم التعليم.

المهارات (بدون إنجازات ولا أرقام هون):
- لو المهارات كلها ورا بعض بدون تقسيم: احكيله يقسّمها لمهارات تخصصية و مهارات شخصية و تحت كل قسم مهاراته. لو أصلاً مقسّمة لمجموعات منظّمة فهي سليمة و تجاوزها بدون ملاحظة.
- لو في أشرطة تقدم أو نجوم أو أيقونات جنب المهارات: احكيله هادا غلط و بيشتت القارئ و ما بيوصل مستواك بشكل واضح و الصح نقاط و تفاصيل صغيرة.
- لو في تكرار بالمهارات (زي ذكر Microsoft Office ثم Word و Excel و PowerPoint كل واحد لحال): احكيله في تكرار و الصح تكتفي بذكر Microsoft Office مرة وحدة أو تفصّل البرامج بدون تكرار. ممنوع تستخدم كلمة "أجنحة" أو "جناح" كترجمة لـ Suite اكتب الاسم بالإنجليزي زي ما هو.

الدورات و الشهادات:
- لأول دورة ناقصة تفاصيل: احكيله عندك دورة (اسم الدورة) مش كاتب (شو الناقص بالضبط) و الصح كل دورة تذكر اسمها + الجهة المانحة يعني مكان الحصول عليها + تاريخ الإنجاز + عدد ساعات الدورة.
- مهم: لو في أكتر من دورة ناقصها نفس الإشي ادمجهم كلهم بملاحظة وحدة و اذكر أسماء الدورات مع بعض بالسطر نفسه. مثال "عندك كمان دورات (اسم و اسم و اسم) مش كاتب عدد ساعاتهم". ممنوع منعاً باتاً تكتب سطر منفصل لكل دورة بنفس النقص لأنو هادا تكرار ممل.
- لو ترتيب الدورات عشوائي: احكيله رتّبهم بشكل أفضل.

اللغات:
- لو ناقص مستوى كل لغة: احكيله مش كاتب مستواك بكل لغة بقسم اللغات و بس. ممنوع تضيف "و الصح توضح المستوى بالكلمات زي Advanced أو Good" الجملة بتنتهي هون.
- لو مقيّم اللغة بنجوم أو أشرطة تقدم: احكيله الصح تكتب المستوى بالكلمات زي Advanced أو Good مش رسومات نجوم ولا نظام أشرطة. لما تذكر الأشرطة اكتبها "نظام أشرطة" مش "أشرطة تقدم".
- لو مستوى اللغات مكتوب صح بالكلمات: قول قسم اللغات ممتاز و تجاوزه. ممنوع تستخدم عبارة "بشكل نصي".

المعرفين (References):
- وجود القسم عادي و مقبول. ممنوع تنصح بحذفه أو تعتبره خطأ.

التصميم و التنسيق:
- لو السيرة أطول من 3 صفحات: احكيله سيرتك طويلة أكثر من اللازم و الصح تكون من صفحتين لـ 3 صفحات. بدون ذكر أي سبب.
- التصميم: ارفضه بس لو القالب مبهرج فعلاً و واضح إنو مش مكتوب على الورد و مش رسمي: يعني خلفية عمود جانبي كامل ملوّنة أو أيقونات أو صورة شخصية أو جداول أو رسومات و أشكال أو ألوان تصميم كتيرة و متنوعة أو قالب جرافيكي ثقيل. بهاي الحالة بس احكيله التصميم مرفوض للأسف لأنو أغلب الشركات بتستخدم برامج فلترة آلية للطلبات فالصح تنكتب السيرة على الورد بقالب بسيط و واضح و رسمي لحتى تتفلتر صح. ممنوع تصف تفاصيل التصميم.
- مهم جداً: القالب اللي فيه عناوين أقسام ملوّنة (أزرق أو أي لون واحد) + عمود جانبي للتواريخ و المكان + خطوط فاصلة رفيعة تحت العناوين هو قالب نظيف بسيط و مكتوب على الورد و مقبول تماماً. تجاوزه بدون أي ملاحظة تصميم إطلاقاً. وجود عمود للتواريخ أو لون للعناوين لحاله مش سبب للرفض.
- لو في مساحات بيضاء كبيرة أو عنوان قسم بصفحة و تفاصيله بصفحة تانية: احكيله هاي المساحات البيضاء و الفراغات خطأ و مش احترافي.
- لو دامج اللغتين العربية و الإنجليزية بنفس السيرة: احكيله دمج اللغتين بنفس السيرة خطأ.
- لو مرفق صور شهادات: احكيله هادول بينحطوا بملف تاني بجانب السيرة عند المقابلة لأنها بتطوّل السيرة و مسؤول التوظيف مش محتاجها بهاي المرحلة."""

CV_USER_PROMPT = "هاي صور سيرتي الذاتية، افحصها واحكيلي ملاحظاتك."


# ---------------------------------------------------------------------------
# التحويل إلى صور
# ---------------------------------------------------------------------------
def pdf_bytes_to_png_list(pdf_data: bytes, dpi: int = ANALYSIS_DPI):
    """يحوّل بايتات PDF إلى قائمة صور PNG (بايتات)."""
    images = []
    with fitz.open(stream=pdf_data, filetype="pdf") as doc:
        for page in doc:
            pix = page.get_pixmap(dpi=dpi)
            images.append(pix.tobytes("png"))
    return images


def find_soffice():
    """يبحث عن مسار LibreOffice (soffice) على النظام، ويعيد None إن لم يُوجد."""
    exe = shutil.which("soffice") or shutil.which("soffice.exe")
    if exe:
        return exe
    candidates = [
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    ]
    for c in candidates:
        if os.path.isfile(c):
            return c
    return None


def word_to_pdf_bytes(word_data: bytes, filename: str) -> bytes:
    """يحوّل ملف Word (بايتات) إلى PDF (بايتات) عبر LibreOffice headless."""
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "LibreOffice غير مثبّت. ثبّته من: https://www.libreoffice.org/download/"
        )

    ext = os.path.splitext(filename)[1] or ".docx"
    with tempfile.TemporaryDirectory() as tmp:
        in_path = os.path.join(tmp, "input" + ext)
        with open(in_path, "wb") as f:
            f.write(word_data)

        profile = os.path.join(tmp, "lo_profile")
        cmd = [
            soffice,
            "--headless", "--norestore", "--nolockcheck",
            f"-env:UserInstallation=file:///{profile.replace(os.sep, '/')}",
            "--convert-to", "pdf",
            "--outdir", tmp,
            in_path,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, timeout=120)

        pdfs = glob.glob(os.path.join(tmp, "*.pdf"))
        if not pdfs:
            raise RuntimeError(
                f"فشل تحويل Word إلى PDF.\n{result.stdout}\n{result.stderr}"
            )
        with open(pdfs[0], "rb") as f:
            return f.read()


# ---------------------------------------------------------------------------
# استدعاء الموديل (xAI Grok — بواجهة متوافقة مع OpenAI)
# ---------------------------------------------------------------------------
_xai_client = None


def get_client() -> AsyncOpenAI:
    """يُنشئ عميل xAI (Grok) مرة واحدة ويعيد استخدامه."""
    global _xai_client
    if _xai_client is None:
        _xai_client = AsyncOpenAI(api_key=XAI_API_KEY, base_url=XAI_BASE_URL)
    return _xai_client


async def analyze_cv_with_ai(images: list[bytes]) -> str:
    """يرسل صور السيرة الذاتية إلى موديل Grok Vision ويعيد نص التقييم."""
    client = get_client()

    content = [{"type": "text", "text": CV_USER_PROMPT}]
    for data in images[:MAX_IMAGES]:
        b64 = base64.b64encode(data).decode("ascii")
        content.append(
            {
                "type": "image_url",
                "image_url": {
                    "url": f"data:image/png;base64,{b64}",
                    "detail": "high",
                },
            }
        )

    resp = await client.chat.completions.create(
        model=XAI_MODEL,
        messages=[
            {"role": "system", "content": CV_SYSTEM_PROMPT},
            {"role": "user", "content": content},
        ],
        temperature=0.6,
        max_tokens=3000,
    )
    return resp.choices[0].message.content


# ---------------------------------------------------------------------------
# منطق المعالجة المشترك
# ---------------------------------------------------------------------------
def normalize_phone(raw: str) -> str | None:
    """ينظّف رقم الواتساب: يشيل المسافات و الشرطات و علامة + و الأقواس.

    "+970 567 785 882" -> "970567785882"
    "0567785882"       -> "970567785882"  (رقم محلي فلسطيني)
    يعيد None إذا كان الرقم غير منطقي.
    """
    digits = re.sub(r"\D", "", raw or "")     # أرقام فقط
    if digits.startswith("00"):
        digits = digits[2:]
    if digits.startswith("0"):
        digits = "970" + digits[1:]
    return digits if 9 <= len(digits) <= 15 else None


async def resolve_group_id(context: ContextTypes.DEFAULT_TYPE) -> int | None:
    """يحدّد صيغة معرّف الجروب الصحيحة (عادي أو سوبر جروب) ويخزّنها."""
    cached = context.application.bot_data.get("group_id")
    if cached:
        return cached

    raw = str(REPORT_GROUP_ID).strip()
    digits = raw.lstrip("-")
    candidates = [raw]
    if not digits.startswith("100"):
        candidates.append("-100" + digits)      # صيغة السوبر جروب

    for cid in candidates:
        try:
            await context.bot.get_chat(int(cid))
            context.application.bot_data["group_id"] = int(cid)
            logger.info("معرّف الجروب الفعّال: %s", cid)
            return int(cid)
        except Exception:
            continue

    logger.error("تعذّر الوصول للجروب %s — تأكد أن البوت عضو فيه", raw)
    return None


def _wa_url(phone: str, report: str, limit: int = WA_URL_LIMIT) -> str:
    """رابط واتساب مع نص التقرير معبّأ (يسقط النص لو طلع أطول من حد الروابط)."""
    full = f"https://wa.me/{phone}?text={quote(report)}"
    if len(full) <= limit:
        return full
    logger.warning("التقرير أطول من حد الرابط — الزر رح يفتح المحادثة بدون نص معبّأ")
    return f"https://wa.me/{phone}"


async def send_status_message(
    context: ContextTypes.DEFAULT_TYPE, group_id: int, phone: str, report: str
):
    """يرسل رسالة الحالة مع زر واحد يفتح واتساب ويعلّم الحالة (تم الإرسال) معاً."""
    wa_url = _wa_url(phone, report)
    text = f"حالة الفحص: ⏳ بانتظار الإرسال\nرقم العميل: +{phone}"

    # بدون رابط عام (تشغيل محلي مثلاً) نرجع لزر التأكيد اليدوي.
    if not PUBLIC_URL:
        markup = InlineKeyboardMarkup(
            [
                [
                    InlineKeyboardButton(
                        "📤 إرسال الفحص للعميل",
                        url=_wa_url(phone, report, WA_URL_LIMIT_BUTTON),
                    )
                ],
                [InlineKeyboardButton("✅ تم الإرسال", callback_data=f"sent:{phone}")],
            ]
        )
        try:
            await context.bot.send_message(group_id, text, reply_markup=markup)
        except Exception:
            logger.exception("تعذّر إرسال رسالة الحالة")
        return

    # الزر يفتح رابطنا الوسيط: يعلّم الحالة ثم يحوّل فوراً إلى واتساب.
    token = secrets.token_urlsafe(8)
    markup = InlineKeyboardMarkup(
        [
            [
                InlineKeyboardButton(
                    "📤 إرسال الفحص للعميل", url=f"{PUBLIC_URL}/s/{token}"
                )
            ]
        ]
    )
    try:
        msg = await context.bot.send_message(group_id, text, reply_markup=markup)
    except Exception:
        logger.exception("تعذّر إرسال رسالة الحالة")
        return

    _PENDING_SENDS[token] = {
        "token": token,
        "chat_id": msg.chat_id,
        "message_id": msg.message_id,
        "phone": phone,
        "wa_url": wa_url,
        "done": False,
    }


async def mark_as_sent(bot, entry: dict):
    """بعد الضغط: زر (تم الإرسال بنجاح) + زر (إرسال الفحص مرة أخرى)."""
    if entry["done"]:
        return
    entry["done"] = True
    phone = entry["phone"]
    try:
        await bot.edit_message_text(
            chat_id=entry["chat_id"],
            message_id=entry["message_id"],
            text=f"حالة الفحص: ✅ تم الإرسال\nرقم العميل: +{phone}",
            reply_markup=InlineKeyboardMarkup(
                [
                    [InlineKeyboardButton("✅ تم الإرسال بنجاح", callback_data="done")],
                    [
                        InlineKeyboardButton(
                            "🔁 إرسال الفحص مرة أخرى",
                            url=f"{PUBLIC_URL}/s/{entry['token']}",
                        )
                    ],
                ]
            ),
        )
    except Exception:
        entry["done"] = False
        logger.exception("تعذّر تحديث حالة الفحص")


async def on_done_clicked(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """زر (تم الإرسال بنجاح) للعرض فقط — يعطي إشعار بدون تغيير الرسالة."""
    await update.callback_query.answer("تم إرسال الفحص للعميل ✅")


async def on_sent_clicked(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """زر التأكيد اليدوي (يُستخدم فقط لما ما يكون في رابط عام)."""
    q = update.callback_query
    await q.answer("تم تحديث الحالة ✅")

    phone = q.data.split(":", 1)[1] if ":" in q.data else ""
    who = q.from_user.first_name or "المستخدم"
    kb = (
        InlineKeyboardMarkup(
            [[InlineKeyboardButton("📤 فتح محادثة العميل", url=f"https://wa.me/{phone}")]]
        )
        if phone
        else None
    )
    try:
        await q.edit_message_text(
            f"حالة الفحص: ✅ تم الإرسال\nرقم العميل: +{phone}\nأرسله: {who}",
            reply_markup=kb,
        )
    except Exception:
        logger.exception("تعذّر تحديث حالة الفحص")


async def process_cv(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    images: list[bytes],
    phone: str | None = None,
):
    """يستقبل قائمة صور جاهزة، يرسلها للموديل، ويعيد التقييم للمستخدم."""
    if not images:
        await context.bot.send_message(chat_id, "❌ لم أستطع قراءة أي صورة من الملف.")
        return

    if not XAI_API_KEY:
        await context.bot.send_message(
            chat_id,
            "❌ مفتاح xAI غير مضبوط. أضف XAI_API_KEY في ملف .env ثم أعد تشغيل البوت.",
        )
        return

    try:
        async with _checks_semaphore():
            report = await analyze_cv_with_ai(images)
        report = sanitize_report(report)

        group_id = await resolve_group_id(context)
        target = group_id or chat_id          # لو تعذّر الجروب نرجع لمحادثة البوت

        # تيليجرام يحدّ الرسالة بـ 4096 حرفاً — نقسّم إن لزم.
        for chunk in _split_text(report, 4000):
            await context.bot.send_message(target, chunk)
        if phone:
            await send_status_message(context, target, phone, report)

        if not group_id:
            await context.bot.send_message(
                chat_id, "⚠️ تعذّر الوصول للجروب — عُرضت النتيجة هنا"
            )
    except Exception as e:
        logger.exception("خطأ أثناء فحص السيرة الذاتية")
        await context.bot.send_message(
            chat_id, f"❌ حدث خطأ أثناء فحص السيرة (+{phone or '—'}):\n{e}"
        )


def _checks_semaphore() -> asyncio.Semaphore:
    """سيمافور مشترك يحدّ عدد الفحوصات المتوازية حتى لا نُغرق واجهة xAI."""
    global _CHECKS_SEM
    if _CHECKS_SEM is None:
        _CHECKS_SEM = asyncio.Semaphore(MAX_CONCURRENT_CHECKS)
    return _CHECKS_SEM


_CHECKS_SEM: asyncio.Semaphore | None = None

# نحتفظ بمراجع مهام الفحص الجارية حتى لا يجمعها الـ GC قبل أن تنتهي.
_RUNNING_CHECKS: set[asyncio.Task] = set()


def schedule_cv_check(
    context: ContextTypes.DEFAULT_TYPE,
    chat_id: int,
    images: list[bytes],
    phone: str | None = None,
):
    """يشغّل الفحص في الخلفية ليقدر المستخدم يرسل عدة سير متتالية بدون انتظار."""
    task = asyncio.create_task(process_cv(context, chat_id, images, phone))
    _RUNNING_CHECKS.add(task)
    task.add_done_callback(_RUNNING_CHECKS.discard)


# حركات التشكيل العربية (فتحة ضمة كسرة شدة سكون تنوين... + الألف الخنجرية).
_TASHKEEL_RE = re.compile(r"[ً-ٰٟ]")


def sanitize_report(text: str) -> str:
    """ينظّف رد الموديل ليطلع كتابة بشرية: بدون تشكيل و لا فواصل و لا نقاط نهاية جمل.

    - يشيل كل حركات التشكيل (شدة كسرة فتحة...).
    - يشيل الفاصلة العربية والإنجليزية.
    - يشيل نقطة نهاية الجملة (النقطة اللي بعدها مسافة أو سطر أو نهاية النص)
      مع الحفاظ على نقط الروابط زي nied.ps لأنها بتكون بين حرفين بدون مسافة.
    """
    if not text:
        return text
    text = _TASHKEEL_RE.sub("", text)          # شيل التشكيل
    text = text.replace("،", "").replace("٫", "")  # فاصلة عربية
    text = re.sub(r"،|,(?=\s|$)", "", text)     # فاصلة إنجليزية بنهاية المقطع
    text = re.sub(r"\.(?=\s|$)", "", text)      # نقطة نهاية جملة (مش نقط الروابط)
    text = re.sub(r"[ \t]+\n", "\n", text)      # مسافات زايدة قبل السطر الجديد
    text = re.sub(r"[ \t]{2,}", " ", text)      # مسافات مزدوجة
    return text.strip()


def _split_text(text: str, size: int):
    """يقسّم النص الطويل إلى أجزاء لا تتجاوز حد رسائل تيليجرام."""
    text = text or "(لا يوجد رد)"
    return [text[i:i + size] for i in range(0, len(text), size)] or [text]


async def images_from_document(doc, context) -> list[bytes]:
    """يحوّل مستند تيليجرام (PDF/Word/صورة) إلى قائمة صور PNG."""
    file_name = doc.file_name or "document"
    lower = file_name.lower()
    tg_file = await doc.get_file()
    file_data = bytes(await tg_file.download_as_bytearray())

    is_pdf = (doc.mime_type == "application/pdf") or lower.endswith(".pdf")
    is_word = lower.endswith(WORD_EXTS)
    is_image = (doc.mime_type or "").startswith("image/") or lower.endswith(
        (".png", ".jpg", ".jpeg", ".webp", ".bmp")
    )

    if is_word:
        file_data = await asyncio.to_thread(word_to_pdf_bytes, file_data, file_name)
        is_pdf = True

    if is_pdf:
        return await asyncio.to_thread(pdf_bytes_to_png_list, file_data)
    if is_image:
        return [file_data]
    return []


# ---------------------------------------------------------------------------
# المعالِجات (Handlers)
# ---------------------------------------------------------------------------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 أهلاً بك في بوت فحص السيرة الذاتية الاحترافي!\n\n"
        "📄 أرسل لي سيرتك الذاتية على شكل:\n"
        "• صورة أو عدة صور (ألبوم)\n"
        "• ملف PDF\n"
        "• ملف Word (docx/doc)\n\n"
        "🔍 وسأفحصها بذكاء وأعطيك تقييماً احترافياً مفصّلاً مع توصيات للتحسين."
    )


async def handle_document(update: Update, context: ContextTypes.DEFAULT_TYPE):
    doc = update.message.document
    if doc.file_size and doc.file_size > MAX_FILE_MB * 1024 * 1024:
        await update.message.reply_text(
            f"❌ حجم الملف كبير جداً (الحد الأقصى {MAX_FILE_MB}MB)."
        )
        return

    chat_id = update.effective_chat.id
    try:
        images = await images_from_document(doc, context)
    except Exception as e:
        logger.exception("خطأ أثناء تحويل الملف")
        await update.message.reply_text(f"❌ تعذّر تحويل الملف:\n{e}")
        return

    if not images:
        await update.message.reply_text(
            "❌ نوع الملف غير مدعوم. أرسل صورة أو PDF أو Word."
        )
        return

    await ask_for_phone(context, chat_id, images)


async def handle_photo(update: Update, context: ContextTypes.DEFAULT_TYPE):
    msg = update.message
    photo = msg.photo[-1]  # أعلى دقة
    tg_file = await photo.get_file()
    data = bytes(await tg_file.download_as_bytearray())

    chat_id = update.effective_chat.id
    mgid = msg.media_group_id

    # صورة مفردة → نطلب رقم العميل ثم نفحص
    if not mgid:
        await ask_for_phone(context, chat_id, [data])
        return

    # ألبوم → نجمّع الصور ثم نعالجها دفعة واحدة (debounce)
    groups = context.application.bot_data.setdefault("media_groups", {})
    grp = groups.setdefault(mgid, {"chat_id": chat_id, "images": [], "task": None})
    grp["images"].append(data)

    if grp["task"]:
        grp["task"].cancel()
    grp["task"] = asyncio.create_task(_flush_media_group(context, mgid))


async def _flush_media_group(context: ContextTypes.DEFAULT_TYPE, mgid: str):
    """ينتظر انتهاء وصول صور الألبوم ثم يعالجها معاً."""
    try:
        await asyncio.sleep(MEDIA_GROUP_DELAY)
    except asyncio.CancelledError:
        return
    groups = context.application.bot_data.get("media_groups", {})
    grp = groups.pop(mgid, None)
    if grp:
        await ask_for_phone(context, grp["chat_id"], grp["images"])


async def ask_for_phone(
    context: ContextTypes.DEFAULT_TYPE, chat_id: int, images: list[bytes]
):
    """يخزّن صور السيرة ويطلب رقم واتساب العميل قبل بدء الفحص."""
    context.application.bot_data.setdefault("pending_cv", {})[chat_id] = images
    await context.bot.send_message(
        chat_id,
        "📱 قبل ما أفحص السيرة ابعتلي رقم الواتس الخاص بالعميل",
    )


async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    pending = context.application.bot_data.setdefault("pending_cv", {})

    # في سيرة منتظرة رقم العميل → هذه الرسالة هي الرقم
    if chat_id in pending:
        phone = normalize_phone(update.message.text)
        if not phone:
            await update.message.reply_text(
                "❌ الرقم مش واضح. ابعت رقم واتساب صحيح\n\nمثال: +970 567 785 882"
            )
            return
        images = pending.pop(chat_id)
        await update.message.reply_text(
            "✅ تم استلام السيرة\nسيتم فحصها وإرسال النتيجة إلى الجروب"
        )
        schedule_cv_check(context, chat_id, images, phone)
        return

    await update.message.reply_text(
        "📄 أرسل لي سيرتك الذاتية (صورة / صور / PDF / Word) لأفحصها لك."
    )


# ---------------------------------------------------------------------------
# الرابط الوسيط: /s/<token> يعلّم الحالة (تم الإرسال) ثم يحوّل إلى واتساب
# ---------------------------------------------------------------------------
# ملاحظة: التخزين بالذاكرة فقط — بعد إعادة تشغيل البوت تفقد الأزرار القديمة صلاحيتها.
_PENDING_SENDS: dict[str, dict] = {}


class _RedirectHandler(BaseHTTPRequestHandler):
    bot = None
    loop: asyncio.AbstractEventLoop | None = None

    def do_GET(self):  # noqa: N802 — اسم مفروض من المكتبة
        path = urlparse(self.path).path

        if not path.startswith("/s/"):
            self._reply(200, b"ok")     # فحص صحة الخدمة (healthcheck)
            return

        entry = _PENDING_SENDS.get(path[3:])
        if not entry:
            self._reply(404, "الرابط منتهي الصلاحية — افتح محادثة العميل يدوياً".encode())
            return

        self.send_response(302)
        self.send_header("Location", entry["wa_url"])
        self.send_header("Cache-Control", "no-store")
        self.end_headers()

        if self.bot and self.loop:
            asyncio.run_coroutine_threadsafe(mark_as_sent(self.bot, entry), self.loop)

    def _reply(self, code: int, body: bytes):
        self.send_response(code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, *args):
        pass                              # نكتم سجلّ الطلبات الافتراضي


async def start_redirect_server(app):
    """يشغّل سيرفر التحويل بخيط جانبي بعد جهوزية البوت."""
    if not PUBLIC_URL:
        logger.warning(
            "PUBLIC_URL غير مضبوط — رح يظهر زر (تم الإرسال) اليدوي بدل التحويل التلقائي"
        )
        return

    _RedirectHandler.bot = app.bot
    _RedirectHandler.loop = asyncio.get_running_loop()
    server = ThreadingHTTPServer(("0.0.0.0", PORT), _RedirectHandler)
    threading.Thread(target=server.serve_forever, daemon=True).start()
    logger.info("سيرفر التحويل يعمل على المنفذ %s (%s)", PORT, PUBLIC_URL)


# ---------------------------------------------------------------------------
# التشغيل
# ---------------------------------------------------------------------------
def main():
    if not BOT_TOKEN:
        raise SystemExit(
            "لم يتم ضبط التوكن. أضف BOT_TOKEN في ملف .env ثم أعد التشغيل."
        )
    if not XAI_API_KEY:
        logger.warning("XAI_API_KEY غير مضبوط — الفحص لن يعمل حتى تضيفه في .env")

    app = (
        Application.builder()
        .token(BOT_TOKEN)
        .post_init(start_redirect_server)
        .build()
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CallbackQueryHandler(on_sent_clicked, pattern=r"^sent:"))
    app.add_handler(CallbackQueryHandler(on_done_clicked, pattern=r"^done$"))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_document))
    app.add_handler(MessageHandler(filters.PHOTO, handle_photo))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))

    logger.info(
        "بوت فحص السيرة الذاتية يعمل الآن (الموديل: %s)... اضغط Ctrl+C للإيقاف.",
        XAI_MODEL,
    )
    app.run_polling(allowed_updates=Update.ALL_TYPES)


if __name__ == "__main__":
    main()
