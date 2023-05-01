import os
import docx
import docx2txt
import language_tool_python
import re

file_text = docx2txt.process("Fuad.docx")
input_file = "Fuad.docx"
output_folder = "D:\Coding\Fuad"


# output_file = "Fuad_corrected.docx"


def process_subtitle_file(input_file, output_file):
    # Check if input file exists and is readable
    if not os.path.isfile(input_file):
        print(f"Input file '{input_file}' does not exist.")
        return

    if not os.access(input_file, os.R_OK):
        print(f"Input file '{input_file}' is not readable.")
        return

    # Import the docx file
    text = docx2txt.process(input_file)

    # Remove timing numbers and empty lines
    lines = text.split('\n')
    clean_lines = []
    for line in lines:
        # Remove lines starting with digits and colon
        if re.match(r'^\d+:\d+', line):
            continue
        if line.strip() == '':
            continue
        clean_lines.append(line)
    text = '\n'.join(clean_lines)

    # Fix grammar mistakes and punctuation marks
    tool = language_tool_python.LanguageTool('de-DE')  #//For German
    #tool = language_tool_python.LanguageTool('en-US') #// for English

    matches = tool.check(text)
    text = tool.correct(text)

    # Save corrected text to a new file
    doc = docx.Document()
    doc.add_paragraph(text)
    doc.save(output_file)

    print(f"Processed file '{output_file}' and saved the corrected version to '{output_file}'.")


# Example usage
process_subtitle_file("Fuad.docx", "Fuad_corrected.docx")
