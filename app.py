from flask import Flask, request, jsonify
from flask_cors import CORS
import openai

app = Flask(__name__)
CORS(app)  # Enable CORS for cross-origin requests

# Initialize OpenAI client
openai.api_key = 'api_keys'  # Ideally, store this in environment variables

def generate_vba_code(user_command):
    """
    Generates VBA code using OpenAI based on the user's command.
    """
    messages = [
        {"role": "system", "content": "You are a helpful assistant that generates clean VBA code for Excel tasks."},
        {"role": "user", "content": f"Generate VBA code to: {user_command}"}
    ]
    response = openai.ChatCompletion.create(
        model="gpt-4",
        messages=messages
    )
    vba_code = response['choices'][0]['message']['content']
    return vba_code.strip()

@app.route('/generate-vba', methods=['POST'])
def generate_vba():
    data = request.get_json()
    user_command = data.get('command')
    if not user_command:
        return jsonify({"error": "No command provided"}), 400

    vba_code = generate_vba_code(user_command)
    return jsonify({"vba_code": vba_code})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
