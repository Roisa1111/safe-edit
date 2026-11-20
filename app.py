
from flask import Flask, render_template, request, send_file
from docx import Document
from docx.shared import Pt
from docx.oxml.ns import qn
from pathlib import Path
from io import BytesIO

app = Flask(__name__)

# 获取项目文件夹位置
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "template.docx"


# 修改 Word 表格单元格
def fill_cell(table, row, col, text):
    cell = table.cell(row, col)
    paragraph = cell.paragraphs[0]
    paragraph.clear()

    run = paragraph.add_run(text)
    run.font.name = "黑体"
    run.font.size = Pt(14)
    run._element.get_or_add_rPr().rFonts.set(
        qn("w:eastAsia"), "黑体"
    )


@app.route("/", methods=["GET", "POST"])
def index():

    if request.method == "POST":

        # 读取用户输入的信息
        data = {
            "name": request.form.get("name", "").strip(),
            "organizer": request.form.get("organizer", "").strip(),
            "category": request.form.get("category", ""),
            "level": request.form.get("level", ""),
            "location": request.form.get("location", "").strip(),
            "time": request.form.get("time", "").strip(),
            "signup_time": request.form.get("signup_time", "").strip(),
            "target": request.form.get("target", "").strip(),
            "number": request.form.get("number", "").strip(),
            "applicant": request.form.get("applicant", "").strip(),
            "phone": request.form.get("phone", "").strip(),
            "content": request.form.get("content", "").strip()
        }

        # 打开原始 Word 模板
        doc = Document(TEMPLATE_PATH)
        table = doc.tables[0]

        # 填写基本信息
        fill_cell(table, 0, 1, data["name"])
        fill_cell(table, 1, 1, data["organizer"])

        # 自动勾选活动类别
        categories = ["思想成长", "文体活动", "创新创业", "技能特长"]
        category_text = "    ".join(
            ("☑ " if item == data["category"] else "□ ") + item
            for item in categories
        )
        fill_cell(table, 2, 1, category_text)

        # 自动勾选活动级别
        levels = ["校级", "院级", "班团支部级"]
        level_text = "    ".join(
            ("☑ " if item == data["level"] else "□ ") + item
            for item in levels
        )
        fill_cell(table, 3, 1, level_text)

        # 填写其他信息
        fill_cell(table, 4, 1, data["location"])
        fill_cell(table, 5, 1, data["time"])
        fill_cell(table, 5, 3, data["signup_time"])
        fill_cell(table, 6, 1, data["target"])
        fill_cell(table, 6, 3, data["number"])
        fill_cell(table, 7, 1, data["applicant"])
        fill_cell(table, 7, 3, data["phone"])
        fill_cell(table, 8, 1, data["content"])

        # 生成 Word 文件
        output = BytesIO()
        doc.save(output)
        output.seek(0)

        return send_file(
            output,
            as_attachment=True,
            download_name="第二课堂活动申报表_已填写.docx",
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    return render_template("index.html")


if __name__ == "__main__":
    app.run(debug=True)
