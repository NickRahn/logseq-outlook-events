from flask import Flask, jsonify
import subprocess
import json

with open('config.json', 'r') as json_file:
    config = json.load(json_file)
util_file = f'{config["utilFileLocation"]}\\util.py'
app = Flask(__name__)

@app.route('/run_script')
def run_script():
    try:
        out = subprocess.Popen(['python', util_file], 
               stdout=subprocess.PIPE, 
               stderr=subprocess.PIPE)
        stdout, stderr = out.communicate()
        if stderr:
            return jsonify({'error': stderr.decode().strip()}), 400
        return stdout, 200
    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(port=5000)