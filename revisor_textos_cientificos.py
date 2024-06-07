import docx
from collections import Counter, defaultdict
import os
from tkinter import Tk, filedialog

def analyze_docx(input_path, output_path):
    # Load the document
    doc = docx.Document(input_path)
    
    # Variables to store results
    font_counter = Counter()
    font_size_counter = Counter()
    indent_counter = Counter()
    indent_specifications = defaultdict(list)
    spacing_counter = Counter()
    line_spacing_details = []
    numbered_lines = []
    alignment_counter = Counter()
    alignment_specifications = defaultdict(list)
    unknown_alignment_texts = []
    left_aligned_texts = []  # For Rule 10
    
    # Get document margins
    sections = doc.sections
    margins = {}
    for section in sections:
        margins['top'] = section.top_margin.cm
        margins['bottom'] = section.bottom_margin.cm
        margins['left'] = section.left_margin.cm
        margins['right'] = section.right_margin.cm

    found_references = False
    
    # Iterate over paragraphs to gather information
    for para in doc.paragraphs:
        # Skip blank lines and lines that only have blank characters
        if para.text.strip() == "":
            continue
        
        if 'REFERÊNCIAS' in para.text or 'BIBLIOGRAFIA' in para.text:
            found_references = True
        
        for run in para.runs:
            # Rule 1: Identify and list font types
            font = run.font
            font_name = font.name
            font_size = font.size.pt if font.size else None
            
            if font_name:
                font_counter[font_name] += 1
                
            # Rule 3: List and quantify font sizes
            if font_size:
                font_size_counter[font_size] += 1
                
            # Rule 5 & 6: Identify indents
            left_indent = para.paragraph_format.left_indent
            if left_indent and left_indent.cm > 1.25 and left_indent.cm != 4:
                indent_value = left_indent.cm
                indent_counter[indent_value] += 1
                indent_specifications[indent_value].append(para.text)
        
        # Rule 7: Identify line spacing greater than 1 cm
        line_spacing = para.paragraph_format.line_spacing
        if line_spacing and line_spacing > 1:
            spacing_counter[line_spacing] += 1
            line_spacing_details.append(para.text)
        
        # Rule 8: Identify lines starting with an integer
        words = para.text.lstrip().split()
        if words and words[0].isdigit():
            alignment_type = get_alignment_type(para.alignment)
            numbered_lines.append((para.text, alignment_type))
        
        # Rule 9: Identify paragraph alignment until "REFERÊNCIAS" or "BIBLIOGRAFIA"
        if not found_references:
            alignment_type = get_alignment_type(para.alignment)
            if alignment_type == 'Unknown':
                unknown_alignment_texts.append(para.text)
            alignment_counter[alignment_type] += 1
            alignment_specifications[alignment_type].append(para.text)
        
        # Rule 10: Identify text with left alignment
        if get_alignment_type(para.alignment) == 'Left':
            left_aligned_texts.append(f'"{para.text}"')
    
    # Rule 2: Quantify by percentages font utilization
    total_fonts = sum(font_counter.values())
    font_percentages = {font: (count / total_fonts) * 100 for font, count in font_counter.items()}
    
    # Rule 4: Quantify the presence of margins (average for all sections)
    margin_percentages = {
        'top': sum(section.top_margin.cm for section in sections) / len(sections),
        'bottom': sum(section.bottom_margin.cm for section in sections) / len(sections),
        'left': sum(section.left_margin.cm for section in sections) / len(sections),
        'right': sum(section.right_margin.cm for section in sections) / len(sections)
    }
    
    # Rule 9: Calculate alignment percentages
    total_alignments = sum(alignment_counter.values())
    alignment_percentages = {align: (count / total_alignments) * 100 for align, count in alignment_counter.items()}
    
    # Write results to the output file
    with open(output_path, 'w') as f:
        f.write("Font Utilization:\n")
        for font, percentage in font_percentages.items():
            f.write(f"{font}: {percentage:.2f}%\n")
        
        f.write("\nFont Sizes Utilization:\n")
        for size, count in font_size_counter.items():
            f.write(f"{size} pt: {count} times\n")
        
        f.write("\nMargins Utilization (average for all sections):\n")
        f.write(f"Top Margin: {margin_percentages['top']:.2f} cm\n")
        f.write(f"Bottom Margin: {margin_percentages['bottom']:.2f} cm\n")
        f.write(f"Left Margin: {margin_percentages['left']:.2f} cm\n")
        f.write(f"Right Margin: {margin_percentages['right']:.2f} cm\n")
        
        f.write("\nIndentations (greater than 1.25 cm and different than 4 cm):\n")
        for indent, count in indent_counter.items():
            f.write(f"{indent} cm: {count} times\n")
        
        f.write("\nSpecific Indents (Different than 4 cm):\n")
        for indent, paragraphs in indent_specifications.items():
            f.write(f"{indent} cm: {len(paragraphs)} times\n")
            for paragraph in paragraphs:
                f.write(f"  - {paragraph[:50]}...\n")  # Print the first 50 characters of each paragraph with specific indent
        
        f.write("\nLine Spacing Greater than 1 cm:\n")
        for line in line_spacing_details:
            f.write(f"  - {line[:50]}...\n")  # Print the first 50 characters of each paragraph with specific line spacing
        
        f.write("\nNumbered Lines and Their Alignment:\n")
        for line, alignment in numbered_lines:
            f.write(f"{line[:50]}... - Alignment: {alignment}\n")
        
        f.write("\nParagraph Alignment Until 'REFERÊNCIAS' or 'BIBLIOGRAFIA':\n")
        for alignment, paragraphs in alignment_specifications.items():
            f.write(f"Alignment: {alignment} - {len(paragraphs)} times ({alignment_percentages[alignment]:.2f}%)\n")
        
        f.write("\nUnknown Alignment Texts:\n")
        for text in unknown_alignment_texts:
            f.write(f"{text[:50]}...\n")  # Print the first 50 characters of each paragraph with unknown alignment

        f.write("\nLeft Aligned Texts:\n")  # For Rule 10
        for text in left_aligned_texts:
            f.write(f"{text}\n")

def get_alignment_type(alignment):
    if alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.LEFT:
        return 'Left'
    elif alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.RIGHT:
        return 'Right'
    elif alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.CENTER:
        return 'Center'
    elif alignment == docx.enum.text.WD_ALIGN_PARAGRAPH.JUSTIFY:
        return 'Justified'
    else:
        return 'Left'

def main():
    root = Tk()
    root.withdraw()  # Hide the root window
    input_dir = filedialog.askdirectory(title="Select Input Directory")
    
    if not input_dir:
        print("No directory selected. Exiting.")
        return

    output_dir = input_dir  # Using the same directory for output

    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)

    # Process each .docx file in the input directory
    for filename in os.listdir(input_dir):
        if filename.endswith('.docx'):
            input_path = os.path.join(input_dir, filename)
            output_path = os.path.join(output_dir, f'{os.path.splitext(filename)[0]}_analysis.txt')
            analyze_docx(input_path, output_path)
            print(f"Analysis of {filename} completed. Results saved to {output_path}")

if __name__ == "__main__":
    main()
