# Developer helper: assembles index.html from src/ (base design CSS + Tabeeby CSS + markup).
import os
root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
src = open(os.path.join(root, "src/index.src.html"), encoding="utf-8").read()
base = open(os.path.join(root, "src/base.css"), encoding="utf-8").read()
extra = open(os.path.join(root, "src/tabeeby.css"), encoding="utf-8").read()
css = "<style>\n" + base.strip() + "\n" + extra.strip() + "\n</style>"
open(os.path.join(root, "index.html"), "w", encoding="utf-8").write(src.replace("/*JAGUAR_CSS*/", css))
print("index.html built")
