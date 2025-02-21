import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PIL import Image
import re
import math
import os

def validate_data(df):
    """Validate the input data format and values."""
    required_columns = ['A/A', 'DBH (cm)', 'Tree height (meters)', 
                       'Form factor (0.4 to 0.6)', 'Cordinates']
    
    # Check for required columns
    missing_columns = [col for col in required_columns if col not in df.columns]
    if missing_columns:
        raise ValueError(f"Missing required columns: {', '.join(missing_columns)}")
    
    # Validate data types and ranges
    if not pd.to_numeric(df['DBH (cm)'], errors='coerce').notnull().all():
        raise ValueError("DBH values must be numeric")
    
    if not pd.to_numeric(df['Tree height (meters)'], errors='coerce').notnull().all():
        raise ValueError("Tree height values must be numeric")
    
    if not pd.to_numeric(df['Form factor (0.4 to 0.6)'], errors='coerce').notnull().all():
        raise ValueError("Form factor values must be numeric")
    
    # Check ranges
    if not ((df['Form factor (0.4 to 0.6)'] >= 0.4) & 
            (df['Form factor (0.4 to 0.6)'] <= 0.6)).all():
        raise ValueError("Form factor must be between 0.4 and 0.6")
    
    if not (df['DBH (cm)'] > 0).all():
        raise ValueError("DBH must be positive")
    
    if not (df['Tree height (meters)'] > 0).all():
        raise ValueError("Tree height must be positive")

def parse_coordinates(coord_str):
    """Parse coordinates in format: Letter then number (e.g., 'A5', 'O13', 'I14')."""
    match = re.match(r'([A-Z]{1,2})(\d+)', coord_str)
    if not match:
        raise ValueError(f"Invalid coordinate format: {coord_str}")
    
    col_str, row = match.groups()
    
    # Convert column string to number (A=1, B=2, ..., AA=27, BB=28, etc.)
    if len(col_str) == 1:
        col = ord(col_str) - ord('A') + 1
    else:
        col = (ord(col_str[0]) - ord('A') + 1) * 26 + (ord(col_str[1]) - ord('A') + 1)
    
    return int(row), col

def calculate_volume(dbh, height, form_factor):
    """Calculate tree volume using the formula V = F × (π/4 × D² × H)."""
    diameter = dbh / 100  # Convert cm to meters
    volume = form_factor * ((math.pi/4) * (diameter ** 2) * height)
    return volume

def create_grid_map(excel_file, background_image=None, x_size=39, y_size=32):
    """Create a grid map of tree volumes."""
    try:
        # Verify excel file exists
        if not os.path.exists(excel_file):
            raise FileNotFoundError(f"Excel file not found: {excel_file}")
        
        # Read Excel file
        print(f"Reading Excel file: {excel_file}")
        df = pd.read_excel(excel_file)
        
        # Validate data format
        validate_data(df)
        
        # Initialize grid
        grid = np.zeros((y_size, x_size))
        count_grid = np.zeros((y_size, x_size))
        
        print("Processing tree data...")
        # Process each tree
        for _, row in df.iterrows():
            try:
                # Parse coordinates
                x, y = parse_coordinates(row['Cordinates'])
                if 1 <= x <= x_size and 1 <= y <= y_size:
                    # Calculate volume
                    volume = calculate_volume(
                        row['DBH (cm)'],
                        row['Tree height (meters)'],
                        row['Form factor (0.4 to 0.6)']
                    )
                    
                    # Add to grid (subtract 1 for 0-based indexing)
                    grid[y-1, x-1] += volume
                    count_grid[y-1, x-1] += 1
            except ValueError as e:
                print(f"Warning: Skipping row {row['A/A']}: {str(e)}")
        
        # Calculate means for cells with multiple trees
        mask = count_grid > 0
        grid[mask] = grid[mask] / count_grid[mask]
        
        # Create visualization
        print("Creating visualization...")
        plt.figure(figsize=(15, 12))
        
        # If background image provided, display it first
        if background_image:
            if not os.path.exists(background_image):
                print(f"Warning: Background image not found: {background_image}")
            else:
                img = Image.open(background_image)
                img = img.resize((x_size * 50, y_size * 50))  # Adjust size to match grid
                plt.imshow(img, extent=[0, x_size, 0, y_size], alpha=0.8)  # Increased background visibility
        
        # Create heatmap with higher transparency
        plt.imshow(grid, cmap='YlOrRd', alpha=0.2)  # Further reduced alpha for more background visibility
        plt.colorbar(label='Tree Volume (m³)')
        
        # Add coordinates
        x_labels = range(1, x_size + 1)
        y_labels = []
        for i in range(y_size):
            if i < 26:
                y_labels.append(chr(65 + i))
            else:
                y_labels.append(chr(65 + ((i-26) // 26)) + chr(65 + ((i-26) % 26)))
        
        plt.xticks(range(x_size), x_labels, rotation=45)
        plt.yticks(range(y_size), y_labels[::-1])
        
        plt.title('Tree Volume Grid Map')
        plt.xlabel('X Coordinate')
        plt.ylabel('Y Coordinate')
        
        # Adjust layout and save
        plt.tight_layout()
        output_file = os.path.join(os.path.dirname(excel_file), 'tree_volume_grid_map.png')
        plt.savefig(output_file, dpi=300, bbox_inches='tight')
        plt.close()
        
        print(f"Grid map has been generated as '{output_file}'")
        
        # Print summary of calculations with added information
        print("\nVolume calculations summary:")
        total_volume = 0
        for _, row in df.iterrows():
            volume = calculate_volume(
                row['DBH (cm)'],
                row['Tree height (meters)'],
                row['Form factor (0.4 to 0.6)']
            )
            total_volume += volume
            print(f"Tree {row['A/A']} at {row['Coordinates']}: "
                  f"DBH={row['DBH (cm)']}cm, Height={row['Tree height (meters)']}m, "
                  f"Volume={volume:.2f}m³")
        
        print(f"\nTotal number of trees: {len(df)}")
        print(f"Total volume: {total_volume:.2f}m³")
        print(f"Average volume per tree: {(total_volume/len(df)):.2f}m³")
        
        return True
        
    except Exception as e:
        print(f"Error: {str(e)}")
        return False

if __name__ == "__main__":
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Define file paths relative to the script directory
    excel_file = os.path.join(script_dir, "tree_data.xlsx")
    background_image = os.path.join(script_dir, "background.png")  # Optional
    
    print("Working directory: d:/timber grid calc/")
    print("Input file: d:/timber grid calc/tree_data.xlsx")
    
    # Create the grid map
    success = create_grid_map(excel_file, background_image)
    
    if not success:
        print("\nPlease ensure:")
        print("1. Your Excel file is named 'tree_data.xlsx' and is in the same directory as this script")
        print("2. The Excel file has the following columns: 'A/A', 'DBH (cm)', 'Tree height (meters)', "
              "'Form factor (0.4 to 0.6)', 'Cordinates'")
        print("3. Form factor values must be between 0.4 and 0.6")
        print("4. All measurements (DBH, height) must be positive numbers")
        print("5. If using a background image, it should be named 'background.png' and be in the same directory")