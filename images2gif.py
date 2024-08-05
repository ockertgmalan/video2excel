from PIL import Image
import os

def images_to_gif(input_folder, output_file, duration=50):
    # Get all image files in the input folder
    image_files = [f for f in os.listdir(input_folder) if f.endswith(('png', 'jpg', 'jpeg', 'bmp', 'gif'))]
    image_files.sort()  # Sort the files to maintain order

    # Open images and append to a list
    images = [Image.open(os.path.join(input_folder, file)) for file in image_files]

    # Save as GIF
    if images:
        images[0].save(
            output_file,
            save_all=True,
            append_images=images[1:],
            duration=duration,
            loop=0
        )
        print(f"GIF saved as {output_file}")
    else:
        print("No images found in the input folder.")

if __name__ == "__main__":
    input_folder = r'G:\Projects\Excel2video\pixel-spreadsheet\stardog_in_excel'
    output_file = r'G:\Projects\Excel2video\pixel-spreadsheet\stardog_in_excel\file.gif'
    images_to_gif(input_folder, output_file)