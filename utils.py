import re
from itertools import groupby
from typing import Iterable

from markitdown import MarkItDown
from pdfminer.high_level import extract_pages
from pdfminer.layout import LTTextContainer, LTTextLineHorizontal
import docx
import logging
import colorlog
from tqdm import tqdm
from unstructured.partition.pdf import partition_pdf


logger = colorlog.getLogger(__name__)
handler = colorlog.StreamHandler()
logger.addHandler(handler)
logger.setLevel(logging.DEBUG)  # Set the desired logging level

formatter = colorlog.ColoredFormatter(
    "%(log_color)s%(levelname)-8s%(reset)s %(message)s",
    log_colors={
        "DEBUG": "cyan",
        "INFO": "green",
        "WARNING": "yellow",
        "ERROR": "red",
        "CRITICAL": "red,bg_white",
    },
)
handler.setFormatter(formatter)


def _font_size(lt_text: LTTextContainer) -> int:
    return int(lt_text.height)


def _join_text(line_group: Iterable[LTTextContainer]) -> str:
    return "\n".join(line.get_text().rstrip() for line in line_group)


def _extract_text_and_font_size(lines: list[LTTextLineHorizontal]):
    """iterate over lines to identify:
    * the most common font size - which we assume to be normal text
    * any line that has a larger font size - which we assume to be a heading
    """

    if len(lines) == 0:
        return {}

    font_sizes = {
        font_size: sum(map(len, line_group))
        for font_size, line_group in groupby(
            sorted(lines, key=_font_size), key=_font_size
        )
    }

    most_common_font_size = max(font_sizes, key=font_sizes.get)

    font_size_mapping = {
        absolute: relative
        for relative, absolute in enumerate(sorted(font_sizes, reverse=True), start=1)
        if absolute > most_common_font_size
    }

    return {
        _join_text(group): font_size_mapping[height]
        for height, group in groupby(lines, key=_font_size)
        if height in font_size_mapping
    }


def convert_pdf(file_path: str) -> str:
    lines = [
        line
        for page in tqdm(list(extract_pages(file_path)), unit="page", leave=False)
        for paragraph in page
        if isinstance(paragraph, LTTextContainer)
        for line in paragraph
        if isinstance(line, LTTextLineHorizontal)
    ]

    headings = _extract_text_and_font_size(lines)

    md = MarkItDown()

    text_content = md.convert(file_path).text_content

    for heading, level in headings.items():
        regex = re.compile(f"^{re.escape(heading)}$", re.MULTILINE)
        replacement = "#" * level + " " + heading.replace("\n", " ")
        text_content = re.sub(regex, replacement, text_content)

    # OCR BACKUP
    if text_content.strip() == "":
        logger.warning(f"No text content found in {file_path}, resorting to OCR")
        elements = partition_pdf(
            filename=file_path, 
            languages=["eng"], 
            verbose=True, 
            strategy="hi_res"
        )

        text_content = ""

        for element in elements:
            text_content += element.text + "\n"
        logger.info(f"Extracted text ({len(text_content)} chars) from {file_path} using OCR")

    return text_content


mapping = {
    "Heading no number": "p",
    "Heading 1": "h1",
    "Heading 2": "h2",
    "Heading 3": "h3",
    "Heading 4": "h4",
    "Heading 5": "h5",
    "Heading 6": "h6",
}


def convert_docx(file_path: str) -> str:
    doc = docx.Document(file_path)

    style_map_dict = {
        paragraph.style.name: mapping[paragraph.style.base_style.name]
        for paragraph in doc.paragraphs
        if paragraph.style.base_style and paragraph.style.base_style.name in mapping
    }

    style_map = "\n".join(
        f"p[style-name='{k}'] => {v}" for k, v in style_map_dict.items()
    )

    md = MarkItDown(style_map=style_map)

    return md.convert(file_path).text_content