"""EPUB 打包、OPF 操作、导航文档生成"""

import os
import re
import html
import shutil
import logging
import zipfile
from datetime import datetime, timezone
from xml.etree import ElementTree as ET

logger = logging.getLogger(__name__)


def extract_epub_template(epub_path, target_dir):
    with zipfile.ZipFile(epub_path, "r") as zf:
        zf.extractall(target_dir)


def pack_epub(source_dir, output_path):
    if os.path.exists(output_path):
        os.remove(output_path)

    mimetype_path = os.path.join(source_dir, "mimetype")
    if not os.path.exists(mimetype_path):
        raise RuntimeError("模板中缺少 mimetype 文件。")

    with open(mimetype_path, "r", encoding="utf-8") as f:
        mimetype_content = f.read().strip()
    if mimetype_content != "application/epub+zip":
        raise RuntimeError(f"模板的 mimetype 文件内容无效：{mimetype_content}")

    with zipfile.ZipFile(output_path, "w") as zf:
        zf.write(mimetype_path, "mimetype", compress_type=zipfile.ZIP_STORED)
        for root_dir, _, files in os.walk(source_dir):
            for filename in files:
                full_path = os.path.join(root_dir, filename)
                rel_path = os.path.relpath(full_path, source_dir).replace("\\", "/")
                if rel_path == "mimetype":
                    continue
                zf.write(full_path, rel_path, compress_type=zipfile.ZIP_DEFLATED)


def cleanup_old_text_files(text_dir):
    for name in os.listdir(text_dir):
        lower_name = name.lower()
        if not lower_name.endswith((".html", ".xhtml")):
            continue
        if lower_name in {"coverpage.xhtml", "nav.xhtml"}:
            continue
        try:
            os.remove(os.path.join(text_dir, name))
        except OSError as e:
            logger.warning(f"无法删除旧文件 {name}: {e}")


def cleanup_old_cover_images(images_dir):
    for name in os.listdir(images_dir):
        lower_name = name.lower()
        if not lower_name.startswith("cover."):
            continue
        try:
            os.remove(os.path.join(images_dir, name))
        except OSError as e:
            logger.warning(f"无法删除旧封面图片 {name}: {e}")
            pass


def update_coverpage(coverpage_path, book_title, cover_src):
    with open(coverpage_path, "r", encoding="utf-8") as f:
        content = f.read()

    escaped_title = html.escape(book_title)

    content, title_count = re.subn(
        r"(<title>).*?(</title>)",
        lambda m: f"{m.group(1)}{escaped_title}{m.group(2)}",
        content,
        count=1,
        flags=re.S,
    )

    if title_count == 0:
        raise RuntimeError("模板中的 coverpage.xhtml 缺少 title 节点。")

    content = re.sub(
        r'(<h1\b[^>]*\btitle=")[^"]*(")',
        lambda m: f"{m.group(1)}{html.escape(book_title, quote=True)}{m.group(2)}",
        content,
        count=1,
        flags=re.S,
    )

    if cover_src:
        escaped_src = html.escape(cover_src, quote=True)
        content, img_count = re.subn(
            r'(<img\b[^>]*\bsrc=")[^"]*(")',
            lambda m: f"{m.group(1)}{escaped_src}{m.group(2)}",
            content,
            count=1,
            flags=re.S,
        )
        if img_count == 0:
            raise RuntimeError("模板中的 coverpage.xhtml 缺少封面图片节点。")
    else:
        content = re.sub(
            r"\s*<img\b[^>]*?/?>\s*", "\n", content, count=1, flags=re.S
        )

    with open(coverpage_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)


def build_nav_tree(documents):
    nodes = []
    stack = []
    for document in documents:
        node = {"document": document, "children": []}
        level = max(1, int(document["level"]))

        while stack and stack[-1]["level"] >= level:
            stack.pop()

        if stack:
            stack[-1]["node"]["children"].append(node)
        else:
            nodes.append(node)

        stack.append({"level": level, "node": node})

    return nodes


def render_nav_nodes(nodes, indent):
    lines = []
    for node in nodes:
        title = html.escape(node["document"]["title"])
        href = html.escape(node["document"]["filename"], quote=True)
        lines.append(f"{indent}<li>")
        lines.append(f'{indent}  <a href="{href}">{title}</a>')
        if node["children"]:
            lines.append(f"{indent}  <ol>")
            lines.extend(render_nav_nodes(node["children"], indent + "    "))
            lines.append(f"{indent}  </ol>")
        lines.append(f"{indent}</li>")
    return lines


def build_nav_html(documents):
    tree_nodes = build_nav_tree(documents)
    lines = [
        "      <li>",
        '        <a href="coverpage.xhtml">封面</a>',
        "      </li>",
    ]
    lines.extend(render_nav_nodes(tree_nodes, indent="      "))
    return "\n".join(lines)


def update_nav(nav_path, documents):
    with open(nav_path, "r", encoding="utf-8") as f:
        content = f.read()

    content = re.sub(
        r"(<title>).*?(</title>)",
        r"\1目录\2",
        content,
        count=1,
        flags=re.S,
    )
    content = re.sub(
        r'(<html\b[^>]*\blang=")[^"]*(")',
        r"\1zh-CN\2",
        content,
        count=1,
        flags=re.S,
    )
    content = re.sub(
        r'(<html\b[^>]*\bxml:lang=")[^"]*(")',
        r"\1zh-CN\2",
        content,
        count=1,
        flags=re.S,
    )
    content = re.sub(
        r"(<link\b[^>]*\bhref=\")[^\"]*(\"[^>]*>)",
        r"\1../Styles/main.css\2",
        content,
        count=1,
        flags=re.S,
    )
    content = re.sub(
        r"(<nav\b[^>]*epub:type=\"toc\"[^>]*>\s*<h1>).*?(</h1>)",
        r"\1目录\2",
        content,
        count=1,
        flags=re.S,
    )

    nav_html = build_nav_html(documents)
    content, toc_count = re.subn(
        r"(<nav\b[^>]*epub:type=\"toc\"[^>]*>.*?<ol>)(.*?)(</ol>)",
        lambda m: f"{m.group(1)}\n{nav_html}\n    {m.group(3)}",
        content,
        count=1,
        flags=re.S,
    )

    if toc_count == 0:
        raise RuntimeError("模板中的 nav.xhtml 缺少目录列表。")

    with open(nav_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)


def set_or_create_dc_text(metadata_node, dc_ns, local_name, text):
    node = metadata_node.find(f"{{{dc_ns}}}{local_name}")
    if node is None:
        node = ET.SubElement(metadata_node, f"{{{dc_ns}}}{local_name}")
    node.text = text
    return node


def xhtml_tag(local_name):
    return f"{{http://www.w3.org/1999/xhtml}}{local_name}"


def local_name(tag):
    return tag.split("}", 1)[-1]


def find_first_by_tag(root, local_name_str):
    for node in root.iter():
        if local_name(node.tag) == local_name_str:
            return node
    return None


def find_child_by_tag(root, local_name_str):
    for node in root:
        if local_name(node.tag) == local_name_str:
            return node
    return None


def normalize_content_opf(opf_path, opf_ns):
    with open(opf_path, "r", encoding="utf-8") as f:
        content = f.read()

    if 'xmlns:opf="' not in content:
        content = content.replace(
            "<metadata>",
            f'<metadata xmlns:opf="{opf_ns}">',
            1,
        )

    content = content.replace("<dc:creator role=", "<dc:creator opf:role=")

    with open(opf_path, "w", encoding="utf-8", newline="\n") as f:
        f.write(content)


def update_content_opf(opf_path, metadata, cover_info, documents):
    opf_ns = "http://www.idpf.org/2007/opf"
    dc_ns = "http://purl.org/dc/elements/1.1/"
    ET.register_namespace("", opf_ns)
    ET.register_namespace("dc", dc_ns)

    tree = ET.parse(opf_path)
    root = tree.getroot()
    metadata_node = root.find(f"{{{opf_ns}}}metadata")
    manifest_node = root.find(f"{{{opf_ns}}}manifest")
    spine_node = root.find(f"{{{opf_ns}}}spine")
    if metadata_node is None or manifest_node is None or spine_node is None:
        raise RuntimeError("模板中的 content.opf 结构不完整。")

    set_or_create_dc_text(metadata_node, dc_ns, "language", "zh-CN")
    set_or_create_dc_text(metadata_node, dc_ns, "title", metadata["title"])
    set_or_create_dc_text(metadata_node, dc_ns, "creator", metadata["author"])
    creator_node = metadata_node.find(f"{{{dc_ns}}}creator")
    if creator_node is not None:
        creator_node.set(f"{{{opf_ns}}}role", "aut")

    for subject_node in list(metadata_node.findall(f"{{{dc_ns}}}subject")):
        metadata_node.remove(subject_node)
    for subject in metadata["subjects"]:
        subject_node = ET.SubElement(metadata_node, f"{{{dc_ns}}}subject")
        subject_node.text = subject

    modified_node = None
    for meta_node in metadata_node.findall(f"{{{opf_ns}}}meta"):
        if meta_node.get("property") == "dcterms:modified":
            modified_node = meta_node
            break
    if modified_node is None:
        modified_node = ET.SubElement(metadata_node, f"{{{opf_ns}}}meta")
        modified_node.set("property", "dcterms:modified")
    modified_node.text = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    cover_item_id = "cover-image"
    nav_id = None
    coverpage_id = None
    chapter_href_pattern = re.compile(r"^Text/\d{4}\.(x?html)$")
    for item_node in list(manifest_node.findall(f"{{{opf_ns}}}item")):
        href = item_node.get("href", "")
        item_id = item_node.get("id", "")
        if chapter_href_pattern.match(href):
            manifest_node.remove(item_node)
            continue
        if re.match(r"^Images/cover\.(jpg|png|webp|bmp)$", href, re.I):
            manifest_node.remove(item_node)
            continue
        if href == "Text/nav.xhtml":
            nav_id = item_id
            item_node.set("href", "Text/nav.xhtml")
            item_node.set("media-type", "application/xhtml+xml")
            item_node.set("properties", "nav")
        if href == "Text/coverpage.xhtml":
            coverpage_id = item_id

    if cover_info:
        cover_item = ET.SubElement(manifest_node, f"{{{opf_ns}}}item")
        cover_item.set("id", cover_item_id)
        cover_item.set("href", cover_info["href"])
        cover_item.set("media-type", cover_info["media_type"])
        cover_item.set("properties", "cover-image")

    chapter_ids = []
    for index, document in enumerate(documents, start=1):
        item_node = ET.SubElement(manifest_node, f"{{{opf_ns}}}item")
        item_id = f"chapter-{index:04d}"
        item_node.set("id", item_id)
        item_node.set("href", f'Text/{document["filename"]}')
        item_node.set("media-type", "application/xhtml+xml")
        chapter_ids.append(item_id)

    for itemref_node in list(spine_node.findall(f"{{{opf_ns}}}itemref")):
        spine_node.remove(itemref_node)

    if coverpage_id:
        cover_ref = ET.SubElement(spine_node, f"{{{opf_ns}}}itemref")
        cover_ref.set("idref", coverpage_id)

    for chapter_id in chapter_ids:
        chapter_ref = ET.SubElement(spine_node, f"{{{opf_ns}}}itemref")
        chapter_ref.set("idref", chapter_id)

    if nav_id:
        nav_ref = ET.SubElement(spine_node, f"{{{opf_ns}}}itemref")
        nav_ref.set("idref", nav_id)
        nav_ref.set("linear", "no")

    tree.write(opf_path, encoding="utf-8", xml_declaration=True)
    normalize_content_opf(opf_path, opf_ns)


def write_epub_contents(root_dir, metadata, cover_info, documents):
    oebps_dir = os.path.join(root_dir, "OEBPS")
    text_dir = os.path.join(oebps_dir, "Text")
    images_dir = os.path.join(oebps_dir, "Images")
    os.makedirs(text_dir, exist_ok=True)
    os.makedirs(images_dir, exist_ok=True)

    cleanup_old_text_files(text_dir)
    cleanup_old_cover_images(images_dir)

    for document in documents:
        out_path = os.path.join(text_dir, document["filename"])
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(document["content"])

    if cover_info:
        cover_output_path = os.path.join(images_dir, cover_info["filename"])
        shutil.copyfile(cover_info["source_path"], cover_output_path)

    coverpage_path = os.path.join(text_dir, "coverpage.xhtml")
    nav_path = os.path.join(text_dir, "nav.xhtml")
    opf_path = os.path.join(oebps_dir, "content.opf")

    update_coverpage(
        coverpage_path, metadata["title"], cover_info["src"] if cover_info else None
    )
    update_nav(nav_path, documents)
    update_content_opf(opf_path, metadata, cover_info, documents)
