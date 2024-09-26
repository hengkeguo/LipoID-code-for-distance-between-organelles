from skimage import io 
from skimage.filters import threshold_otsu 
from skimage.morphology import remove_small_objects 
from skimage.measure import label, regionprops 
import matplotlib.pyplot as plt 
import numpy as np 
import pandas as pd
import os

plt.rcParams['font.sans-serif'] = ['SimHei']

# Add pixel to micron conversion factor
pixel_to_micron = 36.9 / 1024
micron_to_pixel_squared = (1024 / 36.9) ** 2

def process_image(image_path):
    # Load image 
    image = io.imread(image_path) 

    # If the image is multi-channel, we assume the red channel is the first channel 
    if len(image.shape) == 3: 
        image = image[:, :, 0] 
        
    # Apply Otsu's method for binarization 
    thresh = threshold_otsu(image) 
    binary = image > thresh 

    # Remove small objects (noise) 
    cleaned_binary = remove_small_objects(binary, min_size=50) 

    # Label connected regions in the binary image
    labeled_droplets = label(cleaned_binary)
    regions = regionprops(labeled_droplets)

    # Extract the area of each lipid droplet and convert to square microns
    droplet_areas_micron = [region.area / micron_to_pixel_squared for region in regions]
    droplet_count = len(droplet_areas_micron)

    # Calculate total area and average area
    total_area = sum(droplet_areas_micron)
    average_area = total_area / droplet_count if droplet_count > 0 else 0

    return {
        "File Name": os.path.basename(image_path),
        "Number of Lipid Droplets": droplet_count,
        "Total Lipid Droplet Area (µm²)": total_area,
        "Average Lipid Droplet Area (µm²)": average_area,
        "Lipid Droplet Area List (µm²)": droplet_areas_micron
    }

# Specify the folder path containing images
folder_path = r'path'

# Get all tif files in the folder
image_files = [f for f in os.listdir(folder_path) if f.endswith('.tif')]

# Process all images and collect results
results = []
for image_file in image_files:
    image_path = os.path.join(folder_path, image_file)
    result = process_image(image_path)
    results.append(result)

# Convert results to DataFrame
df = pd.DataFrame(results)

# Save results to Excel file
excel_path = os.path.join(folder_path, 'fileName.xlsx')

with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    # Save main results
    df[['File Name', 'Number of Lipid Droplets', 'Total Lipid Droplet Area (µm²)', 'Average Lipid Droplet Area (µm²)']].to_excel(writer, sheet_name='Summary Results', index=False)
    
    # Save detailed lipid droplet area data for each image
    for result in results:
        file_name = result['File Name'].split('.')[0]  # Remove file extension
        droplet_data = pd.DataFrame({
            'Lipid Droplet Number': range(1, len(result['Lipid Droplet Area List (µm²)']) + 1),
            'Lipid Droplet Area (µm²)': result['Lipid Droplet Area List (µm²)']
        })
        droplet_data.to_excel(writer, sheet_name=file_name, index=False)

print(f"Results have been saved to {excel_path}")