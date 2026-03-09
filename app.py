import streamlit as st
import pandas as pd
import requests
import json
import base64
import os
from io import BytesIO
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# --- CONFIGURATION ---
OPENROUTER_API_KEY = os.getenv("OPENROUTER_API_KEY", "")
MODEL_NAME = "openai/gpt-4o"

# --- HELPER FUNCTIONS ---
def encode_image(image_bytes):
    return base64.b64encode(image_bytes).decode('utf-8')

def parse_image_with_openrouter(image_bytes):
    """Sends the image to OpenRouter and returns the structured JSON."""
    if not OPENROUTER_API_KEY:
        st.error("OpenRouter API Key is missing. Please set it in your .env file.")
        return None

    base64_image = encode_image(image_bytes)

    headers = {
        "Authorization": f"Bearer {OPENROUTER_API_KEY}",
        "Content-Type": "application/json"
    }

    prompt = """
    Act as a precise data entry clerk for a wood factory. 
    You will inspect a handwritten or drawn note containing material orders.

    Extract the dimensions and apply these strict factory domain rules:
    1. Dimensions: If written in meters (e.g. 1,00 or 0,45), convert entirely to millimeters (e.g. 1000 or 450).
    2. PVC Tape (Edge Banding): If a dimension is underlined on the note or drawing, it means PVC tape is required on that edge.
    3. Material: Extract any material codes exactly as written (e.g., 326).
    4. Quantity: Extract only the numeric value (e.g., "τεμ 12" becomes 12).

    Output the result STRICTLY as a JSON array of objects mimicking the order rows.
    Do NOT wrap the JSON in markdown formatting (like ```json), just output the raw JSON array.
    Each object MUST have these exact keys:
    [
      {
        "Material": "string (e.g., '326')",
        "Description": "string (e.g., 'door', optional)",
        "Length_mm": integer,
        "Width_mm": integer,
        "Quantity": integer,
        "PVC_Length": boolean (true if length dimension is underlined),
        "PVC_Width": boolean (true if width dimension is underlined)
      }
    ]
    """

    payload = {
        "model": MODEL_NAME,
        "messages": [
            {
                "role": "user",
                "content": [
                    {
                        "type": "text",
                        "text": prompt
                    },
                    {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/jpeg;base64,{base64_image}"
                        }
                    }
                ]
            }
        ]
    }

    try:
        response = requests.post("https://openrouter.ai/api/v1/chat/completions", headers=headers, json=payload)
        response.raise_for_status()
        
        result = response.json()
        content = result['choices'][0]['message']['content'].strip()
        
        # Clean up potential markdown formatting if the model disobeys instructions
        if content.startswith("```json"):
            content = content[7:]
        if content.endswith("```"):
            content = content[:-3]
            
        return json.loads(content)
        
    except requests.exceptions.RequestException as e:
        st.error(f"API Request failed: {e}")
        return None
    except json.JSONDecodeError as e:
        st.error(f"Failed to parse JSON response from the model. Raw response:\n{content}")
        return None
    except Exception as e:
        st.error(f"An unexpected error occurred: {e}")
        return None

def generate_excel(df):
    """Converts the pandas dataframe to an Excel byte stream."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

# --- STREAMLIT UI ---
st.set_page_config(page_title="VTSD Pipeline", layout="wide")

st.title("🏭 Factory Vision-to-Data Pipeline")
st.markdown("Upload a handwritten wood/melamine order to automatically extract it into an editable table using AI.")

# File uploader (allows camera input on mobile/tablets)
uploaded_file = st.file_uploader("Upload or take a picture of the order", type=["jpg", "jpeg", "png"])

if uploaded_file is not None:
    # Display the image
    st.image(uploaded_file, caption="Uploaded Image", use_column_width=True)
    
    if st.button("Extract Data", type="primary"):
        with st.spinner("Analyzing image... This may take 10-15 seconds."):
            image_bytes = uploaded_file.getvalue()
            parsed_json = parse_image_with_openrouter(image_bytes)
            
            if parsed_json:
                st.session_state['parsed_data'] = parsed_json
                st.success("Analysis complete! Please review the extracted data.")

# Display and edit data
if 'parsed_data' in st.session_state:
    st.subheader("Validation (Edit cells to fix errors)")
    
    # Load into dataframe
    df = pd.DataFrame(st.session_state['parsed_data'])
    
    # Render editable dataframe
    edited_df = st.data_editor(
        df,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "PVC_Length": st.column_config.CheckboxColumn("PVC Length?", default=False),
            "PVC_Width": st.column_config.CheckboxColumn("PVC Width?", default=False),
        }
    )
    
    # Generate Excel button
    st.divider()
    excel_data = generate_excel(edited_df)
    st.download_button(
        label="📥 Download Extracted Order as Excel",
        data=excel_data,
        file_name="extracted_order.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
