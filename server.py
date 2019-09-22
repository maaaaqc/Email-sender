import json
from pathlib import Path

from flask import Flask, jsonify, request

import run_all

app = Flask(__name__)


@app.route("/custom", methods=['POST'])
def custom():
    data = request.get_json()
    filename = data['filename']
    to_email = data['email']
    target = data['target']
    return common(filename, to_email, json.loads(target))


@app.route("/default", methods=['POST'])
def default():
    data = request.get_json()
    filename = data['filename']
    to_email = data['email']
    with open(run_all.CONFIG) as json_file:
        target = json.load(json_file)
    return common(filename, to_email, target['target'])


def common(filename, email, target):
    run_all.server_call(filename, target, email)
    return jsonify(target)


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=7710)
