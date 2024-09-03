import os
from flask import Flask, request, jsonify
from flask.cli import load_dotenv
from flask_cors import CORS  
import openai
import xlwings as xw

app = Flask(__name__)
CORS(app)  


client = openai.OpenAI(api_key=api_keys)

def openai_prompt(messages):
    """
    Send a prompt to OpenAI GPT and return the response.
    """
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
        {"role": "system", "content": "You are a helpful assistant that generates clean VBA code for Excel tasks. Do not include any explanations, triple backticks, or 'vba' annotations. Output only the VBA code needed to complete the task. handle type mismatches, ensure empty columns do not cause errors, and allow for expanding beyond the current data range and take into consideration possible edge cases the excel might have depending on the user request.You are also responsible to be able to forecast depending on the user request and using the best method you deem right, moreover if asked you should plot what the user wants and save it as an PNG, make sure to only plot only when asked.Make sure you understand user request perfectly and recognize if he is addressing column header names or column numbers. Make sure the generated code satisfies user request and does not raise errors.Insure the code you generate is excel compatible and enables excel to track formulas, and use any of its built it capabilities. do not explain just generate the code."},
        {"role": "user", "content": f"Generate VBA code to: {user_command}"}
    ]
    vba_code = openai_prompt(messages)
    clean_vba_code = vba_code.replace('```', '').replace('vba', '').strip()
    return clean_vba_code

def run_vba_code(wb, vba_code):
    try:
        # Add VBA code into a new module
        vba_module = wb.api.VBProject.VBComponents.Add(1)  # 1 is for standard module
        vba_module.CodeModule.AddFromString(vba_code)

        # Ensure the subroutine is named properly
        sub_name = vba_code.split("Sub ")[1].split("(")[0].strip()

        # Fully qualify the macro name with the workbook name
        macro_name = f"'{wb.name}'!{sub_name}"

        # Run the VBA macro using Application.Run
        wb.api.Application.Run(macro_name)

        # Optionally, remove the module after running to keep the workbook clean
        # vba_module.Delete()

        return "VBA Code executed successfully!"
    except Exception as e:
        # Print additional debug information
        import traceback
        error_details = traceback.format_exc()
        return f"Error executing VBA code: {str(e)}\nDetails: {error_details}"


@app.route('/run-command', methods=['POST'])
def run_command():
    data = request.get_json()
    user_command = data['command']

    # Get the active workbook
    wb = xw.books.active

    # Generate VBA code using OpenAI
    vba_code = generate_vba_code(user_command)

    # Run the generated VBA code in the active workbook
    result = run_vba_code(wb, vba_code)

    return jsonify({"message": result, "vba_code": vba_code})

if __name__ == '__main__':
    app.run(port=5000, debug=True)