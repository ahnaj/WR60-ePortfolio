import os
import pypandoc

# Ensure pandoc is installed
try:
    pypandoc.get_pandoc_path()
except OSError:
    pypandoc.download_pandoc()

# Define source and output folders
folder = "./Documents"
output_folder = "./PDF/"
os.makedirs(output_folder, exist_ok=True)

# Convert all .docx files in the folder to .pdf
for filename in os.listdir(folder):
    if filename.endswith(".docx"):
        input_path = os.path.join(folder, filename)
        output_filename = os.path.splitext(filename)[0] + ".pdf"
        output_path = os.path.join(output_folder, output_filename)

        print(f"Converting {filename} → {output_filename}")
        try:
            pypandoc.convert_file(input_path, 'pdf', outputfile=output_path)
        except Exception as e:
            print(f"❌ Failed to convert {filename}: {e}")

print("✅ All possible conversions completed.")
