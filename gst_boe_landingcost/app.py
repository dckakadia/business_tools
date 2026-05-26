import os
from flask import Flask, render_template

app = Flask(__name__)
app.secret_key = os.environ.get('PORTAL_SECRET', 'gw-portal-s3cr3t-k3y-2026')
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'output'
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Register Blueprints
from blueprints.gst_reconciliation import gst_bp
from blueprints.boe_calculator import boe_bp
from blueprints.landing_cost import landing_bp

app.register_blueprint(gst_bp)
app.register_blueprint(boe_bp)
app.register_blueprint(landing_bp)

class ScriptNameStripper(object):
    def __init__(self, app):
        self.app = app
    def __call__(self, environ, start_response):
        script_name = environ.get('HTTP_X_SCRIPT_NAME', '')
        if script_name:
            environ['SCRIPT_NAME'] = script_name.rstrip('/')
        return self.app(environ, start_response)

app.wsgi_app = ScriptNameStripper(app.wsgi_app)

@app.route('/')
def dashboard():
    return render_template('dashboard.html')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5001, debug=True)
