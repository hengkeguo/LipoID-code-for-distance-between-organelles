from skimage import io 
from skimage.filters import threshold_otsu 
from scipy.ndimage import distance_transform_edt 
from skimage.measure import label, regionprops
import numpy as np 
import matplotlib.pyplot as plt 
import pandas as pd
import os

plt.rcParams['font.sans-serif'] = ['SimHei']

# Add pixel to micron conversion factor
pixel_to_micron = 36.9 / 1024

def process_image(image_path):
    # Load image 
    image = io.imread(image_path) 
    # Separate channels 
    red_channel = image[:, :, 0] # Lipid droplets (red) 
    blue_channel = image[:, :, 1] # Microtubules (blue) 
    green_channel = image[:, :, 2] # Mitochondria (green) 
    
    # Use Otsu's method for binarization 
    thresh_red = threshold_otsu(red_channel) 
    binary_red = red_channel > thresh_red 
    thresh_blue = threshold_otsu(blue_channel) 
    binary_blue = blue_channel > thresh_blue 
    thresh_green = threshold_otsu(green_channel) 
    binary_green = green_channel > thresh_green 

    # Calculate distance transform for microtubules
    distance_to_blue = distance_transform_edt(~binary_blue) * pixel_to_micron

    # Label lipid droplets and mitochondria
    labeled_red = label(binary_red)
    labeled_green = label(binary_green)

    # Get distances for each lipid droplet and mitochondrion
    red_regions = regionprops(labeled_red)
    green_regions = regionprops(labeled_green)

    red_distances = [distance_to_blue[prop.coords[:, 0], prop.coords[:, 1]].min() for prop in red_regions]
    green_distances = [distance_to_blue[prop.coords[:, 0], prop.coords[:, 1]].min() for prop in green_regions]

    # Calculate summary statistics
    min_distance_red_to_blue = min(red_distances) if red_distances else np.inf
    min_distance_green_to_blue = min(green_distances) if green_distances else np.inf
    mean_distance_red_to_blue = np.mean(red_distances) if red_distances else np.nan
    mean_distance_green_to_blue = np.mean(green_distances) if green_distances else np.nan

    return {
        "Minimum distance (Lipid to Microtubule)": min_distance_red_to_blue if min_distance_red_to_blue != np.inf else "Lipid area is empty",
        "Minimum distance (Mitochondria to Microtubule)": min_distance_green_to_blue if min_distance_green_to_blue != np.inf else "Mitochondria area is empty",
        "Average distance (Lipid to Microtubule)": mean_distance_red_to_blue if not np.isnan(mean_distance_red_to_blue) else "Lipid area is empty",
        "Average distance (Mitochondria to Microtubule)": mean_distance_green_to_blue if not np.isnan(mean_distance_green_to_blue) else "Mitochondria area is empty",
        "Lipid to Microtubule distance list": red_distances,
        "Mitochondria to Microtubule distance list": green_distances
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
    result['File name'] = image_file
    results.append(result)

# Convert results to DataFrame
df = pd.DataFrame(results)

# Save results to Excel file
excel_path = os.path.join(folder_path, 'fileName.xlsx')

with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
    # Save summary results
    df[['File name', 'Minimum distance (Lipid to Microtubule)', 'Minimum distance (Mitochondria to Microtubule)', 
        'Average distance (Lipid to Microtubule)', 'Average distance (Mitochondria to Microtubule)']].to_excel(writer, sheet_name='Summary Results', index=False)
    
    # Save detailed distance data for each image
    for result in results:
        file_name = result['File name'].split('.')[0]  # Remove file extension
        red_data = pd.DataFrame({
            'Lipid Number': range(1, len(result['Lipid to Microtubule distance list']) + 1),
            'Lipid to Microtubule Distance (microns)': result['Lipid to Microtubule distance list']
        })
        green_data = pd.DataFrame({
            'Mitochondria Number': range(1, len(result['Mitochondria to Microtubule distance list']) + 1),
            'Mitochondria to Microtubule Distance (microns)': result['Mitochondria to Microtubule distance list']
        })
        combined_data = pd.concat([red_data, green_data], axis=1)
        combined_data.to_excel(writer, sheet_name=file_name, index=False)

print(f"Results have been saved to {excel_path}")