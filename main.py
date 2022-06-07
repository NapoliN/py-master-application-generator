from ast import arg
import docx
import unicodedata
import sys

# 行当たりの幅（適当）
ROW_WIDTH = 94

args = sys.argv
docx_file = args[1]
your_file = args[2]

# カス形式のdocxファイルを開く
doc = docx.Document(docx_file)

tables = doc.tables
not_write, name_and_choice, body1, body2 = tables[0], tables[1], tables[2], tables[3]

# 志望理由書を開く
my_text = open(your_file,encoding="UTF-8")

# 名前と希望系を入力
name = my_text.readline()
name_and_choice.rows[0].cells[1].text = name
choice = my_text.readline()
name_and_choice.rows[0].cells[3].text = choice

# 本文を入力
txt = my_text.read()
my_text.close()

txt_ptr = 0
for body in [body1,body2]:
    for row in body.rows:
        injection = ""
        now_width = 0
        while now_width < ROW_WIDTH:
            if txt_ptr >= len(txt):
                print("Successfully Done!")
                doc.save("result.docx")
                exit(0)
            c = txt[txt_ptr]
            c_width = 2 if unicodedata.east_asian_width(c) in "FWA" else 1
            if c == "\n":
                txt_ptr += 1
                break
            if c_width + now_width > ROW_WIDTH:
                break
            txt_ptr += 1
            now_width += c_width
            injection += c
        row.cells[0].text = injection

print("Number of characters limit error")
doc.save("result.docx")