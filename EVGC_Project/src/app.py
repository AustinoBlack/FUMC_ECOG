from flask import Flask, render_template, request, redirect, url_for, send_from_directory, session
from werkzeug.utils import secure_filename
import os
import shutil
import atexit
import secrets
from evgp_cli import process_pptx, get_upcoming_sunday, generate_preview

app = Flask(__name__)
app.secret_key = secrets.token_hex(16)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.config['PREVIEW_FOLDER'] = 'preview'

os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['PREVIEW_FOLDER'], exist_ok=True)

ALLOWED_EXTENSIONS = {'pptx', 'jpg', 'jpeg', 'png', 'gif'}

def allowed_file(filename):
    ext = filename.rsplit('.', 1)[-1].lower()
    return '.' in filename and ext in ALLOWED_EXTENSIONS and '/' not in filename

def cleanup_directories():
    for dir_path in ["preview", "outputs", "uploads"]:
        abs_path = os.path.join(os.getcwd(), dir_path)
        if os.path.exists(abs_path):
            for file_name in os.listdir(abs_path):
                file_path = os.path.join(abs_path, file_name)
                if os.path.isfile(file_path) or os.path.islink(file_path):
                    os.unlink(file_path)
                elif os.path.isdir(file_path):
                    shutil.rmtree(file_path)
    print("Cleanup complete.")

atexit.register(cleanup_directories)

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    preview_images = []
    preview_folder = None
    sunday_folder = get_upcoming_sunday()

    uploaded_filename = session.get('filename')

    if request.method == 'POST':
        file = request.files.get('file')
        icon = request.files.get('icon')
        background_color = request.form.get('bg_color')
        font_name = request.form.get('font_family')

        # Handle file upload or reuse
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            input_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(input_path)
            session['pptx_path'] = input_path
            session['filename'] = filename
        else:
            input_path = session.get('pptx_path')
            filename = session.get('filename')

        if not input_path or not os.path.exists(input_path):
            return "Please upload a PowerPoint file to continue.", 400

        # Save icon
        icon_path = None
        if icon and allowed_file(icon.filename):
            icon_filename = secure_filename(icon.filename)
            icon_path = os.path.join(app.config['UPLOAD_FOLDER'], icon_filename)
            icon.save(icon_path)

        # Background handling
        background_choice = request.form.getlist('background')
        is_image_bg = 0
        bg_image_path = None

        bg_image = request.files.get('bg_image')
        if 'Image' in background_choice and bg_image and allowed_file(bg_image.filename):
            is_image_bg = 1
            bg_image_filename = secure_filename(bg_image.filename)
            bg_image_path = os.path.join(app.config['UPLOAD_FOLDER'], bg_image_filename)
            bg_image.save(bg_image_path)
        elif 'Color' in background_choice:
            is_image_bg = 0
            bg_image_path = None

        output_folder = os.path.join(app.config['OUTPUT_FOLDER'], sunday_folder)
        os.makedirs(output_folder, exist_ok=True)

        # Process PowerPoint
        process_pptx(input_path, output_folder, icon_path, background_color, font_name, is_image_bg, bg_image_path)

        # Generate preview
        output_pptx_path = os.path.join(output_folder, "output.pptx")
        preview_path = os.path.join(app.config['PREVIEW_FOLDER'], filename.rsplit('.', 1)[0])
        preview_images = generate_preview(output_pptx_path, preview_path)
        preview_folder = os.path.basename(preview_path)

        return render_template('index.html',
                               download=True,
                               sunday_folder=sunday_folder,
                               preview_images=preview_images,
                               preview_folder=preview_folder,
                               uploaded_filename=filename)

    return render_template('index.html', download=False, uploaded_filename=uploaded_filename)

@app.route('/download')
def download_file():
    sunday_folder = request.args.get('sunday_folder')
    if sunday_folder:
        return send_from_directory(os.path.join(app.config['OUTPUT_FOLDER'], sunday_folder), 'output.pptx', as_attachment=True)
    return redirect(url_for('upload_file'))

@app.route('/preview/<path:filename>')
def preview_file(filename):
    return send_from_directory(app.config['PREVIEW_FOLDER'], filename)

@app.route('/clear')
def clear_session():
    session.clear()
    return redirect(url_for('upload_file'))

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

