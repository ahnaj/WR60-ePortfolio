import os
import pypandoc

# Ensure pandoc is installed
try:
    pypandoc.get_pandoc_path()
except OSError:
    pypandoc.download_pandoc()

# Define source folder
folder = "./Documents"  # Use current working directory or update path

# Convert all .docx files in the folder to .html
for filename in os.listdir(folder):
    if filename.endswith(".docx"):
        input_path = os.path.join(folder, filename)
        output_filename = os.path.splitext(filename)[0] + ".html"
        output_path = os.path.join("./HTML/", output_filename)

        print(f"Converting {filename} to {output_filename}")
        pypandoc.convert_file(input_path, 'html', outputfile=output_path)

print("All conversions completed.")
