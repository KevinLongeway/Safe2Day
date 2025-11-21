"""
select_logo.py - Logo selection interface
Displays available logos from Client_Logos folder and stores user selection.
"""

import os
from pathlib import Path
import json


def get_available_logos():
    """
    Scan Client_Logos folder and return list of available logo files.
    
    Returns:
        List of logo file paths
    """
    logos_folder = Path(__file__).parent / "Client_Logos"
    
    if not logos_folder.exists():
        print(f"Error: Client_Logos folder not found at {logos_folder}")
        return []
    
    # Look for common image formats
    image_extensions = ['.png', '.jpg', '.jpeg', '.bmp', '.gif', '.tiff']
    logos = []
    
    for ext in image_extensions:
        logos.extend(logos_folder.glob(f"*{ext}"))
    
    return sorted(logos)


def display_logo_menu(logos):
    """
    Display available logos and get user selection.
    
    Args:
        logos: List of logo file paths
        
    Returns:
        Selected logo path or None
    """
    print("\n" + "=" * 60)
    print("Available Client Logos:")
    print("=" * 60)
    
    for idx, logo in enumerate(logos, 1):
        print(f"{idx}. {logo.name}")
    
    print(f"{len(logos) + 1}. Cancel")
    print("=" * 60)
    
    while True:
        try:
            choice = input("\nSelect a logo number: ").strip()
            choice_num = int(choice)
            
            if choice_num == len(logos) + 1:
                print("Selection cancelled.")
                return None
            
            if 1 <= choice_num <= len(logos):
                selected_logo = logos[choice_num - 1]
                print(f"\n✓ Selected: {selected_logo.name}")
                return selected_logo
            else:
                print(f"Please enter a number between 1 and {len(logos) + 1}")
        
        except ValueError:
            print("Please enter a valid number")
        except KeyboardInterrupt:
            print("\n\nSelection cancelled.")
            return None


def save_logo_selection(logo_path):
    """
    Save the selected logo path to a JSON file.
    
    Args:
        logo_path: Path to the selected logo
    """
    selection_file = Path(__file__).parent / "selected_logo.json"
    
    selection_data = {
        "logo_path": str(logo_path),
        "logo_name": logo_path.name
    }
    
    with open(selection_file, 'w') as f:
        json.dump(selection_data, f, indent=2)
    
    print(f"✓ Selection saved to: {selection_file}")


def select_logo():
    """
    Main function to handle logo selection process.
    
    Returns:
        Path to selected logo or None
    """
    logos = get_available_logos()
    
    if not logos:
        print("No logo files found in Client_Logos folder.")
        print("Please add logo files (.png, .jpg, .jpeg, .bmp, .gif, .tiff)")
        return None
    
    selected_logo = display_logo_menu(logos)
    
    if selected_logo:
        save_logo_selection(selected_logo)
        return selected_logo
    
    return None


if __name__ == "__main__":
    print("=" * 60)
    print("Safe2Day Logo Selection Tool")
    print("=" * 60)
    
    result = select_logo()
    
    if result:
        print("\n✓ Logo selection complete!")
    else:
        print("\nNo logo selected.")
