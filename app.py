from flask import Flask, request, jsonify, send_file
import json, os, subprocess, tempfile, sys

app = Flask(__name__)

@app.route('/generate', methods=['POST'])
def generate():
    data = request.get_json()
    rows = data.get('rows', [])
    date_label = data.get('date_label', 'This Period')
    api_key = data.get('api_key', os.environ.get('ANTHROPIC_API_KEY'))
    
    with tempfile.NamedTemporaryFile(mode='w', suffix='.json', delete=False) as f:
        json.dump(rows, f)
        data_path = f.name
    
    output_path = tempfile.mktemp(suffix='.pptx')
    
    result = subprocess.run([
        sys.executable, 'generate_aunz_deck.py',
        '--data', data_path,
        '--output', output_path,
        '--date-label', date_label,
        '--api-key', api_key
    ], capture_output=True, text=True)
    
    os.unlink(data_path)
    
    if result.returncode != 0:
        return jsonify({'error': result.stderr}), 500
    
    return send_file(output_path, as_attachment=True, 
                     download_name='aunz_refund_report.pptx')

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)))
