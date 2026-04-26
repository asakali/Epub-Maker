"""HTML/XHTML 模板渲染"""

import re
import html
import os


def render_group(template, title):
    if "{{TITLE}}" in template:
        return template.replace("{{TITLE}}", html.escape(title))

    result = template

    result = re.sub(
        r"(<title>).*?(</title>)",
        rf"\1{html.escape(title)}\2",
        result,
        count=1,
        flags=re.S,
    )
    result = re.sub(
        r'(<h1\b[^>]*title=")[^"]*(")',
        rf"\1{html.escape(title)}\2",
        result,
        count=1,
        flags=re.S,
    )
    result = re.sub(
        r"(<h1\b[^>]*></h1>\s*)(<p\b[^>]*class=\"k1\"[^>]*>).*?(</p>)",
        rf"\1\2{html.escape(title)}\3",
        result,
        count=1,
        flags=re.S,
    )
    return result


def render_content(
    content_template,
    full_title,
    heading_tag,
    chapter_no,
    chapter_subtitle,
    paragraphs,
):
    content_html = "\n".join(
        f'  <p class="lth">{html.escape(p)}</p>' for p in paragraphs
    )

    if "{{TITLE}}" in content_template:
        result = content_template
        result = result.replace("{{TITLE}}", html.escape(full_title))
        result = result.replace("{{CHAPTER_NO}}", html.escape(chapter_no))
        result = result.replace(
            "{{CHAPTER_SUBTITLE}}", html.escape(chapter_subtitle)
        )
        result = result.replace("{{CONTENT}}", content_html)
        if heading_tag != "h2":
            result = result.replace("<h2 ", f"<{heading_tag} ", 1)
            result = result.replace("</h2>", f"</{heading_tag}>", 1)
        return result

    result = content_template

    result = re.sub(
        r"(<title>).*?(</title>)",
        rf"\1{html.escape(full_title)}\2",
        result,
        count=1,
        flags=re.S,
    )

    if heading_tag != "h2":
        result = re.sub(r"<h2\b", f"<{heading_tag}", result, count=1)
        result = re.sub(r"</h2>", f"</{heading_tag}>", result, count=1)

    result = re.sub(
        rf'(<{heading_tag}\b[^>]*title=")[^"]*(")',
        rf"\1{html.escape(full_title)}\2",
        result,
        count=1,
        flags=re.S,
    )

    chapter_head_html = (
        f'<p class="k2"><span class="k2h">{html.escape(chapter_no)}</span>'
        f"<br/>{html.escape(chapter_subtitle)}</p>"
        if chapter_subtitle
        else f'<p class="k2"><span class="k2h">{html.escape(chapter_no)}</span></p>'
    )

    result = re.sub(
        r'<p\b[^>]*class="k2"[^>]*>.*?</p>',
        chapter_head_html,
        result,
        count=1,
        flags=re.S,
    )

    body_match = re.search(
        r"(<p\b[^>]*class=\"k2\"[^>]*>.*?</p>)(.*?)(</body>)", result, flags=re.S
    )
    if body_match:
        middle = body_match.group(2)

        img_match = re.search(
            r"(<br\s*/?><img\b[^>]*?/?>.*?<br\s*/?><br\s*/?>)", middle, flags=re.S
        )
        img_block = img_match.group(1) if img_match else "\n"

        replacement = (
            body_match.group(1)
            + img_block
            + "\n"
            + content_html
            + "\n"
            + body_match.group(3)
        )
        result = (
            result[: body_match.start()] + replacement + result[body_match.end():]
        )

    return result


def build_rendered_documents(
    sections, group_template, content_template, extension
):
    documents = []
    for section in sections:
        item = section["item"]
        filename = f'{item["seq"]}.{extension}'
        if section["is_group"]:
            content = render_group(group_template, item["title"])
        else:
            content = render_content(
                content_template=content_template,
                full_title=item["title"],
                heading_tag=f'h{min(item["level"], 6)}',
                chapter_no=item["prefix"] or section["title_line"],
                chapter_subtitle=item["subtitle"] or "",
                paragraphs=section["paragraphs"],
            )

        documents.append(
            {
                "filename": filename,
                "content": content,
                "level": item["level"],
                "title": item["title"],
            }
        )
    return documents


def get_cover_info(cover_image_path):
    if not cover_image_path:
        return None

    source_ext = os.path.splitext(cover_image_path)[1].lower()
    ext_map = {
        ".jpg": (".jpg", "image/jpeg"),
        ".jpeg": (".jpg", "image/jpeg"),
        ".png": (".png", "image/png"),
        ".webp": (".webp", "image/webp"),
        ".bmp": (".bmp", "image/bmp"),
    }
    if source_ext not in ext_map:
        return None

    output_ext, media_type = ext_map[source_ext]
    filename = f"cover{output_ext}"
    return {
        "source_path": cover_image_path,
        "filename": filename,
        "href": f"Images/{filename}",
        "src": f"../Images/{filename}",
        "media_type": media_type,
    }
