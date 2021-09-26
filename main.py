from docx import Document
import docx2txt
import os

# Passing docx file to process function
path_to_file = '/home/danny/Downloads/Документ.docx'
text = docx2txt.process(path_to_file)

# Saving content inside docx file into output.txt file
with open("output.txt", "w") as text_file:
    print(text, file=text_file)

# Math logic and operations
with open('output.txt') as f:
    lines = []
    for line in f:
        pure_line = line.strip('\n')
        lines.append(pure_line)
        while "" in lines:
            lines.remove("")

    """
    Math operations here
    """
    sum0_1 = int(lines[0]) + int(lines[1])
    dev2_01 = int(lines[2]) - sum0_1
    f.close()

print(lines)


# Creating result.txt file to save calculations
file = open("result.txt", "w")
file.write(str(sum0_1) + '\n')
file.write(str(dev2_01) + '\n')
file.close()


doc = Document()


with open("result.txt", 'r', encoding='utf-8') as openfile:
    line = openfile.read()
    doc.add_paragraph(line)
    doc.save("result.docx")
    # doc.save("Your wanted location" + "result.docx") - save directly to preferable directory


# Deleting intermediary files
def remover():
    if os.path.exists("result.txt"):
        os.remove("result.txt")
    else:
        pass

    if os.path.exists("output.txt"):
        os.remove("output.txt")
    else:
        pass


remover()
