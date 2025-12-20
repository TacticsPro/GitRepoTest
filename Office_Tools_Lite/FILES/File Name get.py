import os
import glob

# Specify the directory where the files are located (current directory in this case)
directory = "."

# Use a wildcard to match all files
files = glob.glob(os.path.join(directory, "*.*"))

# Print filenames in one line with quotation marks
print(", ".join(f'"{os.path.basename(file)}"' for file in files))
