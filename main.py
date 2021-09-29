from docx import Document
import docx2txt
import os
import re

# Passing docx file to process function
path_to_file = 'Initial data.docx'
text = docx2txt.process(path_to_file)

# Saving content inside docx file into output.txt file
with open("output.txt", "w") as text_file:
    print(text, file=text_file)

# Preparing our values from strings for future calculations
with open('output.txt') as f:
    lines = []
    values = []
    final_list = []
    for line in f:
        pure_line = line.strip('\n')
        lines.append(pure_line)
        while "" in lines:
            lines.remove("")
    lines.pop(0)
    for element in lines:  # getting only numbers from string
        values += re.findall(r"[-+]?\d*\.\d+|\d+", element)
    for element in values:
        if "." not in element:  # getting integers and floats
            final_list.append(int(element))
        else:
            final_list.append(float(element))
    """
    Setting variables to our values from .docx
    """
    a = final_list[0]  # value a
    b = final_list[1]  # value b
    c = final_list[2]  # value c
    # etc
    """
        Math operations here
    """
    d = (a + b) / c
    d_val = " N, J, kgf"  # specify dimension
    f.close()


# Creating result.txt file to save calculations
file = open("result.txt", "w")
file.write("Results:\n" + '\n')
file.write("d = " + str(d) + d_val + '\n')
file.close()


doc = Document()


with open("result.txt", 'r', encoding='utf-8') as openfile:
    line = openfile.read()
    doc.add_paragraph(line)
    doc.save("Results of calculations.docx")
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
