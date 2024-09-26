import numpy as np
from skimage.io import imread
from skimage.filters import threshold_otsu
from skimage.measure import label, regionprops
from scipy.ndimage import distance_transform_edt
import pandas as pd
import os

# Conversion factor: pixels to micrometers
pixel_to_um = 36.9 / 1024

def detailed_analysis(image):
    # Extract red and green channels
    re_channel = image[:, :, 0]  # Red channel corresponds to lipid droplets
    green_channel = image[:, :, 1]  # Green channel corresponds to mitochondria

    # Apply Otsu thresholding to both channels
    re_thresh = threshold_otsu(re_channel)
    re_binary = re_channel > re_thresh
    green_thresh = threshold_otsu(green_channel)
    green_binary = green_channel > green_thresh

    # Label thresholded regions
    labeled_re = label(re_binary)
    labeled_green = label(green_binary)

    # Calculate distance transform for mitochondria (green channel)
    distance = distance_transform_edt(~green_binary)

    # Extract properties of labeled regions
    re_properties = regionprops(labeled_re)

    # Calculate distances from lipid droplet edges to nearest mitochondria edges
    distances = []
    for re_prop in re_properties:
        # Get coordinates of all pixels in the lipid droplet region
        coords = re_prop.coords
        # Calculate the minimum distance from any pixel in the lipid droplet to the nearest mitochondria
        min_distance = np.min(distance[coords[:, 0], coords[:, 1]])
        distances.append(min_distance * pixel_to_um)

    return distances

def process_folder(folder_path):
    # Get all tif files in the folder
    image_paths = [f for f in os.listdir(folder_path) if f.endswith('.tif')]

    all_data = []
    for path in image_paths:
        full_path = os.path.join(folder_path, path)
        img = imread(full_path)
        name = path.split(".")[0]  # Extract image name from the path
        distances = detailed_analysis(img)
        df = pd.DataFrame({name: distances})
        all_data.append(df)

    # Combine all data into a single DataFrame
    all_data_df = pd.concat(all_data, axis=1)

    # Save results to Excel file
    excel_path = os.path.join(folder_path, "fileName.xlsx")
    with pd.ExcelWriter(excel_path) as writer:
        all_data_df.to_excel(writer, sheet_name='All Data', index=False)
    
    print(f"Results have been saved to {excel_path}")

if __name__ == '__main__':
    # Please enter the path of the folder you want to process here
    folder_to_process = r"path"
    process_folder(folder_to_process)