import os
from flask import Flask, request, jsonify
from flask.cli import load_dotenv
from flask_cors import CORS
import openai

app = Flask(__name__)
CORS(app)

client = openai.OpenAI(api_key=api_keys)

def openai_prompt(messages):
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=messages
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        return str(e)

def generate_vba_code(user_command):
    messages = [
        {"role": "system", "content": "You are a helpful assistant that generates clean VBA code for Excel tasks. Do not include any explanations, triple backticks, or 'vba' annotations. Output only the VBA code needed to complete the task."},
        {"role": "user", "content": f"Generate VBA code to: {user_command}"}
    ]
    vba_code = openai_prompt(messages)
    clean_vba_code = vba_code.replace('```', '').replace('vba', '').strip()
    return clean_vba_code

@app.route('/generate-vba', methods=['POST'])
def generate_vba():
    data = request.get_json()
    user_command = data['command']

    # Generate VBA code using OpenAI
    vba_code = generate_vba_code(user_command)

    return jsonify({"vba_code": vba_code})

if __name__ == '__main__':
    app.run(port=5000, debug=True)
