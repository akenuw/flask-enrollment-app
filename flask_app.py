from flask import Flask, render_template, send_file
import pandas as pd
import os

app = Flask(__name__)
file_name = "employee_data_with_web_access.xlsx"

@app.route("/")
def home():
    try:
        # Read the Excel file
        if not os.path.exists(file_name):
            return "<h1>No database file found. Please create some entries first!</h1>"

        data = pd.read_excel(file_name)
        data_html = data.to_html(classes="table table-bordered table-hover table-striped", index=False)
        
        return f"""
        <!DOCTYPE html>
        <html>
        <head>
            <title>Admin Dashboard</title>
            <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.1.3/dist/css/bootstrap.min.css">
            <style>
                body {{
                    background-color: #f8f9fa;
                }}
                .container {{
                    margin-top: 50px;
                    padding: 20px;
                    background: #ffffff;
                    box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.1);
                    border-radius: 8px;
                }}
                h1 {{
                    font-size: 2rem;
                    margin-bottom: 20px;
                }}
                .btn {{
                    margin-right: 10px;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Employee Database</h1>
                {data_html}
                <div class="mt-3">
                    <a href="/download" class="btn btn-primary">Download Database</a>
                    <a href="/refresh" class="btn btn-secondary">Refresh</a>
                </div>
            </div>
        </body>
        </html>
        """
    except Exception as e:
        return f"<h1>Error reading database: {e}</h1>"

@app.route("/download")
def download():
    try:
        return send_file(file_name, as_attachment=True)
    except Exception as e:
        return f"Error: {e}"

@app.route("/refresh")
def refresh():
    # Simply redirect back to the home page to refresh the table
    return home()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000)
