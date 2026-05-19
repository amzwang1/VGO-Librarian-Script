from datetime import datetime, timezone
import io
import re
from typing import Callable

from PyPDF2 import PdfReader
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase.pdfmetrics import stringWidth
from reportlab.pdfgen import canvas

from score import Binder

def get_user_input(prompt, validator=None):
    # type: (str, Callable[[str], bool]|None) -> str
    print(prompt, end='')
    response = input()
    if validator is None:
        return response
    while not validator(response):
        print("Invalid input, please re-enter: ", end='')
        response = input()
    return response

def yes_no_validator(input):
    return input.lower() in ['y', 'n', 'yes', 'no']

def is_file_name_sanitized(name):
    # type: (str) -> bool
    return re.match(r'^[^\\/:*?"<>|]*$', name) is not None

def yellow_text(text: str):
    return f'\033[93;1m{text}\033[00m'


def get_max_font_size(text, page_width, font_name="Helvetica", margin=54):
    test_size = 12
    text_width = stringWidth(text, font_name, test_size)
    available_width = page_width - 2 * margin
    return test_size * available_width / max(text_width, 1)

class CoverTemplate:
    def __init__(self, pagesize=letter):
        self.w, self.h = pagesize
        # Pre-load and cache the image
        self.logo = ImageReader('assets/vgo_logo.png')

    def generate_pdf(self, binder, date):
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)

        title_font_size = min(64, get_max_font_size(binder.title, self.w))
        c.setFontSize(title_font_size)
        c.drawCentredString(0.5*self.w, 0.85*self.h, binder.title)

        name_font_size = min(32, get_max_font_size(binder.names, self.w))
        c.setFontSize(name_font_size)
        c.drawCentredString(0.5*self.w, 0.15*self.h, binder.names)

        c.setFontSize(16)
        c.drawCentredString(0.5*self.w, 0.07*self.h, date.strftime('%Y-%m-%d'))

        c.drawImage(self.logo, 0.1*self.w, 0.25*self.h, width=0.8*self.w, height=0.5*self.h, preserveAspectRatio=True)

        c.showPage()
        c.showPage()
        c.save()
        buffer.seek(0)
        return PdfReader(buffer)

cover_template = CoverTemplate()

def make_cover_page(binder, date):
    # type: (Binder, datetime) -> PdfReader
    return cover_template.generate_pdf(binder, date)

def to_utc(dt):
    # type: (datetime) -> datetime
    ts = dt.timestamp()
    return datetime.fromtimestamp(ts, tz=timezone.utc)
