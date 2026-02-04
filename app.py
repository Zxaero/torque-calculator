import os
from flask import Flask, jsonify

app = Flask(__name__)

@app.route("/")
def home():
    return """
    <h2>Bolt Torque Calculator</h2>
    <p>✅ App is running on Render</p>
    <p>Next step: add torque logic + SharePoint connection</p>
    """

@app.route("/health")
def health():
    return jsonify(status="ok")

if __name__ == "__main__":
    app.run()
