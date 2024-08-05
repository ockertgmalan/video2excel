import cv2
import os

def extract_scenes(video_path, output_folder, num_scenes):
    # Create the output folder if it doesn't exist
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)

    # Open the video file
    cap = cv2.VideoCapture(video_path)
    if not cap.isOpened():
        print(f"Error: Could not open video file {video_path}")
        return

    # Get the total number of frames in the video
    total_frames = int(cap.get(cv2.CAP_PROP_FRAME_COUNT))

    # Calculate the interval between frames to capture
    interval = total_frames // num_scenes
    
    num_digits = len(str(num_scenes))

    # Iterate through the video frames and save the selected frames
    for i in range(num_scenes):
        frame_number = i * interval
        cap.set(cv2.CAP_PROP_POS_FRAMES, frame_number)
        ret, frame = cap.read()
        if ret:
            output_path = os.path.join(output_folder, f"scene_{i+1:0{num_digits}}.jpg")
            cv2.imwrite(output_path, frame)
        else:
            print(f"Error: Could not read frame {frame_number}")

    # Release the video capture object
    cap.release()

if __name__ == "__main__":
    video_path = r'G:\Downloads\perfect strangers reel.mp4'
    output_folder = r'G:\Projects\Excel2video\pixel-spreadsheet\stardog_frames'
    num_scenes = 1600  # Define the number of scenes to extract
    extract_scenes(video_path, output_folder, num_scenes)