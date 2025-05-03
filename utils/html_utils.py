# utils/html_utils.py
from bs4 import BeautifulSoup

def process_html_content(paragraph, html_content):
    """
    Traite les balises HTML et ajoute du texte formaté dans un paragraphe
    en utilisant BeautifulSoup.
    """
    soup = BeautifulSoup(html_content, "html.parser")
    for element in soup:
        if isinstance(element, str):
            run = paragraph.add_run(element)
            run.font.name = 'Arial'
        elif element.name == 'i':
            run = paragraph.add_run(element.get_text())
            run.font.name = 'Arial'
            run.italic = True
        elif element.name == 'br':
            run = paragraph.add_run()
            run.add_break()
            run.add_break()