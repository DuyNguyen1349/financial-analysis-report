
from flask import Flask, render_template, send_file
import pandas as pd

app = Flask(__name__)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/recommendations")
def recommendations():
    df = pd.read_csv("data/ai_investment_recommendation.csv")
    table_html = df.to_html(index=False, classes="table", border=1)
    return render_template("recommendations.html", table=table_html)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
