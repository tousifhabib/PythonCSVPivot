import argparse
import io
import json
import zipfile
from pathlib import Path
from flask import Flask, request, jsonify, send_file
import logging

from config import config_handling
from data_processing import data_processing_controller

logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(module)s:%(message)s')
logger = logging.getLogger(__name__)

app = Flask(__name__)

DEFAULT_CONFIG_FILE = Path('config/config.json')


@app.route('/generate-files', methods=['POST'])
def generate_files():
    config = config_handling.read_config(DEFAULT_CONFIG_FILE)

    json_data = request.get_json()
    if not json_data:
        return jsonify({'error': 'No JSON data provided'}), 400

    pdf_file, excel_file = data_processing_controller.process_json_data_and_generate_files(config, json_data)

    if pdf_file and excel_file:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w') as zip_file:
            zip_file.write(pdf_file, arcname=Path(pdf_file).name)
            zip_file.write(excel_file, arcname=Path(excel_file).name)

        zip_buffer.seek(0)

        return send_file(
            zip_buffer,
            mimetype='application/zip',
            as_attachment=True,
            download_name='output_files.zip'
        )
    else:
        return jsonify({'error': 'Failed to generate files'}), 500


def run_api():
    app.run()


def process_data(args: argparse.Namespace):
    config_file = Path(args.config_file) if args.config_file else DEFAULT_CONFIG_FILE
    config = config_handling.read_config(config_file) if not args.interactive else {
        'data': config_handling.get_interactive_config(),
        'styles': config_handling.read_config(config_file).get('styles', {})
    }

    if args.json_data:
        try:
            json_data = json.loads(args.json_data)
            data_processing_controller.process_json_data_and_generate_files(config, json_data)
        except json.JSONDecodeError as e:
            logger.error(f"Error parsing JSON data: {e}")
    else:
        data_processing_controller.process_data_and_generate_files(config)


def main():
    parser = argparse.ArgumentParser(description="Build PDF and Excel files.")
    parser.add_argument('--config-file', help="Path to the configuration file")
    parser.add_argument('--interactive', action='store_true', help="Run in interactive mode")
    parser.add_argument('--json-data', help="JSON data to process")
    parser.add_argument('--api', action='store_true', help="Run in API mode")

    args = parser.parse_args()

    if args.api:
        run_api()
    else:
        process_data(args)


if __name__ == '__main__':
    main()
