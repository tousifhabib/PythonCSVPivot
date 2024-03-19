import argparse

from flask import Flask, request, jsonify

from config.config_handling import read_config, get_interactive_config
from data_processing.data_processing_controller import process_data_and_generate_files, \
    process_json_data_and_generate_files
import logging

logging.basicConfig(level=logging.INFO, format='%(levelname)s:%(module)s:%(message)s')

DEFAULT_CONFIG_FILE = 'config/config.json'

app = Flask(__name__)


@app.route('/generate-files', methods=['POST'])
def generate_files():
    config_file = DEFAULT_CONFIG_FILE
    config = read_config(config_file)

    json_data = request.get_json()
    if not json_data:
        return jsonify({'error': 'No JSON data provided'}), 400

    pdf_file, excel_file = process_json_data_and_generate_files(config, json_data)

    return jsonify({
        'pdf_file': pdf_file,
        'excel_file': excel_file
    })


def main(args: argparse.Namespace) -> None:
    if args.api:
        app.run()
    else:
        config_file = args.config_file or DEFAULT_CONFIG_FILE
        config = read_config(config_file) if not args.interactive else {'data': get_interactive_config(),
                                                                        'styles': read_config(config_file).get('styles',
                                                                                                               {})}
        if args.json_data:
            process_json_data_and_generate_files(config, args.json_data)
        else:
            process_data_and_generate_files(config)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Build PDF and Excel files.")
    parser.add_argument('--config-file', help="Path to the configuration file")
    parser.add_argument('--interactive', action='store_true', help="Run in interactive mode")
    parser.add_argument('--json-data', help="JSON data to process")
    parser.add_argument('--api', action='store_true', help="Run in API mode")

    main(parser.parse_args())
