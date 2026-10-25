import os
import sys

# Test reading the template dataset
dataset_path = r"C:\Users\OLUWAFEMI JOSHUA\Desktop\Economics Indicators.xlsx"

if not os.path.exists(dataset_path):
    print("Dataset not found at:", dataset_path)
    sys.exit(1)

print("Found dataset at:", dataset_path)
