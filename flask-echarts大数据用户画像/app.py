from flask import Flask
from jinja2 import Markup
import data_explore
app = Flask(__name__, static_folder="templates")

@app.route('/')
def index():
    c=data_explore.grid_vertical()
    return Markup(c.render_embed())


if __name__ == '__main__':
    app.run()
