# Industrial Lumber VTSD Pipeline

Automated translation of handwritten melamine and wood orders into structured industrial Excel formats using AI Vision.

## Architecture

1. **Capture**: A Streamlit web interface for factory workers to upload photos of handwritten notes via PC/tablet.
2. **Vision Parsing (OpenRouter/GPT-4o)**: Translates handwriting using factory-specific logic (e.g., converting "1,00" meters to 1000mm, identifying underlines as PVC edge instructions, mapping material codes).
3. **Structured Validation**: Presents the AI's transcription in an editable Streamlit dataframe.
4. **Excel Generation**: Finalizes the validated data into the expected industrial `.xlsx` template.

## Worker's User Manual

### How to use the app:
1. Open the tool on your tablet or PC.
2. Click "Upload Image" and take a clear photo of the customer's note.
3. Wait about 10-15 seconds for the AI to interpret the handwriting.
4. **Important:** Compare the results in the table to the original image! Make sure:
    - Dimensions are in millimeters.
    - PVC properties (`Yes` if the dimension was underlined) are correct.
    - The correct material colors or codes are mapped.
5. If there are mistakes, edit the table directly by clicking on the cell.
6. Click "Generate Excel" to download the final order sheet!

### 📸 Tips for Perfect Photos (For Factory Staff)
- **Lighting is King**: Take the photo in a well-lit area. Avoid shadows across the paper.
- **Top-Down Angle**: Try to take the photo from directly above the paper (not angled), so all numbers are easily readable.
- **Keep it Flat**: A flat page is easier to read than a crumpled or folded one.
- **Clarity over Speed**: Ensure the camera has focused on the text before snapping the shot. Blurry numbers look like different numbers.

## Setup & Local Installation

### Requirements
- Python 3.10+
- `pip`
- An OpenRouter API key

### Steps
1. Clone the repository.
2. Create a virtual environment:
   ```bash
   python -m venv .venv
   source .venv/bin/activate
   ```
3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
4. Access OpenRouter API: Create a `.env` file and set your key:
   ```env
   OPENROUTER_API_KEY="sk-or-v1-..."
   ```
5. Run the Streamlit app:
   ```bash
   streamlit run app.py
   ```

## Deployment
This app can be deployed easily via Render or Railway using the included settings. It requires minimal server setup; just link your GitHub account and inject your environment variables into the platform securely.
