from flask import Flask, render_template, request
import os
from topsis import topsis   # correct import

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def index():

    if request.method == 'POST':

        try:
            file = request.files['file']
            weights = request.form['weights']
            impacts = request.form['impacts']

            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            output = os.path.join(UPLOAD_FOLDER, "result.xlsx")

            topsis(filepath, weights, impacts, output)

            return "TOPSIS completed successfully"

        except Exception as e:
            return str(e)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
