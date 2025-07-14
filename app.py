
import requests
from flask import Flask, request, jsonify
from pptx import Presentation
from io import BytesIO

app = Flask(__name__)

@app.route('/extract-text', methods=['POST'])
def extract_ppt_text():
    """
    Receives a PPT file URI via POST request, downloads the file,
    extracts text content from slides, and returns it.
    """
    data = request.get_json()

    if not data or 'ppt_uri' not in data:
        return jsonify({"error": "Invalid input: 'ppt_uri' is required in the JSON body."}), 400

    ppt_uri = data['ppt_uri']

    try:
        # Download the PPT file content
        response = requests.get(ppt_uri, stream=True, verify=False)
        response.raise_for_status() # Raise an exception for bad status codes (4xx or 5xx)

        # Use BytesIO to treat the downloaded content as a file in memory
        ppt_file_in_memory = BytesIO(response.content)

        # Open the presentation
        prs = Presentation(ppt_file_in_memory)

        extracted_text = []

        # Iterate through slides
        for slide_number, slide in enumerate(prs.slides):
            slide_text = f"--- Slide {slide_number + 1} ---\n"
            slide_content = []

            # Iterate through shapes in the slide
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    slide_content.append(shape.text.strip())

            # Join all non-empty text parts from shapes
            if slide_content:
                slide_text += "\n".join(filter(None, slide_content))
            else:
                slide_text += "No text on this slide."

            extracted_text.append(slide_text)

        # Combine text from all slides
        formatted_text = "\n\n".join(extracted_text)

        return jsonify({"success": True, "text_content": formatted_text})

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"Failed to download PPT file: {e}"}), 500
    except Exception as e:
        return jsonify({"error": f"Failed to process PPT file: {e}"}), 500

# Note: We don't need the if __name__ == '__main__': block for Vercel deployment
# Vercel handles the server execution.

if __name__ == '__main__':
    # You can change the host and port as needed
    app.run(host='0.0.0.0', port=5000, debug=True)