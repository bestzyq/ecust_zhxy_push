import os
import zipfile
import shutil
from docx import Document
from docx.shared import Pt
from docx.oxml import ns
from docx.oxml import OxmlElement

def export_images_to_text(docx_file):
    document = Document(docx_file)
    image_counter = 1
    image_folder = "images"  # 图片保存的文件夹名

    # 创建保存图片的文件夹
    current_dir = os.path.dirname(docx_file)
    image_folder_path = os.path.join(current_dir, image_folder)
    os.makedirs(image_folder_path, exist_ok=True)

    for i, element in enumerate(document.inline_shapes):
        if element.type == 3:  # InlineShapeType.PICTURE
            run = element._inline.graphic.graphicData.pic.blipFill.blip
            image_rel_id = run.embed
            image_part = document.part.related_parts[image_rel_id]
            image_data = image_part.blob

            # 保存图片文件
            image_path = os.path.join(image_folder_path, "图片{}.png".format(image_counter))
            with open(image_path, "wb") as f:
                f.write(image_data)

            # 替换图片为文本
            image_name = "图片{}".format(image_counter)
            element.text = "【{}】".format(image_name)

            # 创建新的run并设置样式
            p = element._inline.getparent().getparent()
            new_run = OxmlElement("w:r")
            new_text = OxmlElement("w:t")
            new_text.text = "【{}】".format(image_name)
            new_run.append(new_text)
            rpr = new_run.get_or_add_rPr()
            color_element = OxmlElement("w:color")
            color_element.set(ns.qn("w:val"), "FF0000")
            rpr.append(color_element)
            p.append(new_run)

            # 删除原始图片
            parent = element._inline.getparent()
            parent.remove(element._inline)

            image_counter += 1

    # 保存修改后的文档
    output_file = "【推送】" + os.path.basename(docx_file)
    output_path = os.path.join(current_dir, output_file)
    document.save(output_path)
    print("图片导出完成，已保存为{}".format(output_path))

    # 创建ZIP文件并压缩文档和图片文件夹
    zip_file = os.path.join(current_dir, os.path.splitext(output_file)[0] + ".zip")
    with zipfile.ZipFile(zip_file, "w") as zipf:
        zipf.write(output_path, os.path.basename(output_path))
        for root, dirs, files in os.walk(image_folder_path):
            for file in files:
                file_path = os.path.join(root, file)
                zipf.write(file_path, os.path.relpath(file_path, current_dir))

    print("ZIP文件创建完成，已保存为{}".format(zip_file))

    # 删除图片文件夹和文档
    os.remove(output_path)
    shutil.rmtree(image_folder_path)

# 自动搜索并转换所有的docx文件
current_dir = os.path.dirname(__file__)
for filename in os.listdir(current_dir):
    if filename.endswith('.docx'):
        docx_file = os.path.join(current_dir, filename)
        print("正在处理文件: ", filename)
        export_images_to_text(docx_file)

print("所有文件处理完成")
