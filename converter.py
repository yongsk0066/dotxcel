from flask import Flask, render_template, request,send_file
from dotxel import paint_exel
import hashlib
import os
import datetime

app = Flask(__name__)

@app.route("/", methods=['GET'])
def hello():
    return render_template('main.html')
    
@app.route('/', methods=['POST'])
def post():
    img = request.files['image']
    if not img:
        error='upload image'
        return render_template('main.html', error=error)
    k = int(request.form['color'])
    size = int(request.form['size'])
    img_name = hashlib.md5(str(datetime.datetime.now()).encode('utf-8')).hexdigest()
    img_path = os.path.join('static/img', img_name + os.path.splitext(img.filename)[-1])
    # result_path = os.path.join('static/results', img_name + '.png')
    img.save(img_path)

    paint_exel(img_path, K=k, size=size)

    os.remove(img_path)

    file_name = f"output.xlsx"
    return send_file(file_name,
                     attachment_filename='dotxcel.xlsx',
                     as_attachment=True)

host_addr = "localhost"
port_num = "8080"

if __name__ == '__main__':
    app.run()