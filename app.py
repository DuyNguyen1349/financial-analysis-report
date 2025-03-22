
from flask import Flask, render_template, send_file
import pandas as pd

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/recommendations")
def recommendations():
    df = pd.read_csv("data/ai_investment_recommendation.csv")
    html_table = df.to_html(index=False)
    return f"<h1>Gợi ý đầu tư AI</h1>{html_table}<br><a href='/'>← Về trang chủ</a>"

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
