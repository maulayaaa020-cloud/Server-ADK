import zipfile
from lxml import etree

path = r"D:\Freelaces\Server\htdocs\adk\test_files\paket3\cover 2 dimulai dari iii\hasil\Docx 5_p3.docx"
with zipfile.ZipFile(path) as z:
    xml = z.read("word/footer1.xml")
print(xml.decode("utf-8"))
