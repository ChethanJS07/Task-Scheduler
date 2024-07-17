# Task Scheduler using Python

This Python script automates the process of opening a PowerPoint presentation, starting the slideshow mode, and ending it after a specified duration using Python's `pyautogui` library.

## Requirements

- Python 3.x
- `pyautogui` library

## Installation

1. Clone the repository or download the script directly.

2. Install the required Python packages:
   ```bash
   pip install pyautogui

## Usage

1. Place your PowerPoint (`*.pptx`) file in the same directory as this script or specify its path in the script.

2. Modify the script variables:
   - `ppt_file`: Path to your PowerPoint presentation.
   - `duration`: Duration in seconds for how long the slideshow should run.

3. Run the script:
   ```bash
   python ppt_slideshow_controller.py

## How It Works

This Python script controls a PowerPoint presentation, automating the process of starting the slideshow and ending it after a specified duration.

1. **Initialization**:
   - The script initializes by importing necessary libraries (`os` and `time`) and setting up variables:
     - `ppt_file`: Path to the PowerPoint presentation (`*.pptx`) file.
     - `duration`: Duration in seconds for how long the slideshow should run.

2. **Starting Slideshow**:
   - It uses `os.startfile()` to open the PowerPoint presentation specified by `ppt_file`.
   - After a brief delay (adjustable via `time.sleep()`), it simulates pressing the F5 key to start the slideshow mode (`Slideshow from beginning`).

3. **Running Slideshow**:
   - The script waits for the duration specified in `duration` using `time.sleep()` while the slideshow runs automatically in PowerPoint.

4. **Ending Slideshow**:
   - After the specified duration, the script simulates pressing the Esc key to exit the slideshow mode and return to normal PowerPoint interface.

5. **Completion**:
   - Once the slideshow ends, the script completes execution.

### Example Scenario

- Suppose `ppt_file` is set to `presentation.pptx` and `duration` is set to `60` seconds.
- The script will open `presentation.pptx`, start the slideshow, wait for 60 seconds, then exit the slideshow mode automatically.

### Adjustments

- You can adjust the `ppt_file` path to point to your specific PowerPoint file.
- Modify `duration` to control how long the slideshow runs before automatically exiting.

