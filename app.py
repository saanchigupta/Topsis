from flask import Flask, render_template, request, jsonify, send_file
import pandas as pd
import numpy as np
import os
import smtplib
from email.message import EmailMessage
import re
from datetime import datetime
from dotenv import load_dotenv
import io

load_dotenv()

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

def read_file(file):
    """Read CSV or Excel file and return DataFrame"""
    filename = file.filename.lower()
    
    try:
        if filename.endswith('.csv'):
            df = pd.read_csv(file)
            return df, None
        elif filename.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(file)
            return df, None
        else:
            return None, "File must be CSV or Excel (.csv, .xlsx, .xls)"
    except Exception as e:
        return None, str(e)

def validate_email(email):
    """Validate email format"""
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(pattern, email) is not None

def validate_weights(weights_str):
    """Validate and parse weights"""
    if not weights_str or not weights_str.strip():
        return None, "Weights are required"
    
    try:
        weights_str = weights_str.strip()
        
        # Check if comma-separated
        if ',' not in weights_str:
            return None, "Weights must be comma-separated (e.g., 1,2,3)"
        
        weights = []
        for w in weights_str.split(","):
            w = w.strip()
            if not w:
                return None, "Empty weight value found. Ensure no consecutive commas"
            try:
                weight = float(w)
                if weight <= 0:
                    return None, f"All weights must be positive numbers. Got {weight}"
                weights.append(weight)
            except ValueError:
                return None, f"'{w}' is not a valid number in weights"
        
        return weights, None
    except Exception as e:
        return None, f"Error parsing weights: {str(e)}"

def validate_impacts(impacts_str):
    """Validate and parse impacts"""
    if not impacts_str or not impacts_str.strip():
        return None, "Impacts are required"
    
    try:
        impacts_str = impacts_str.strip()
        
        # Check if comma-separated
        if ',' not in impacts_str:
            return None, "Impacts must be comma-separated (e.g., +,-,+)"
        
        impacts = []
        for i in impacts_str.split(","):
            i = i.strip()
            if not i:
                return None, "Empty impact value found. Ensure no consecutive commas"
            if i not in ['+', '-']:
                return None, f"Each impact must be '+' or '-'. Got '{i}'"
            impacts.append(i)
        
        return impacts, None
    except Exception as e:
        return None, f"Error parsing impacts: {str(e)}"

def validate_csv_file(df):
    """Validate CSV file structure and content"""
    # Check minimum columns
    if df.shape[1] < 3:
        return None, f"CSV must contain at least 3 columns (1 name + at least 2 criteria). Found {df.shape[1]} columns"
    
    # Check numeric columns (from 2nd to last)
    non_numeric_issues = []
    
    # Start from column index 1 (2nd column) to the last column
    for col_idx in range(1, df.shape[1]):
        col_name = df.columns[col_idx]
        try:
            # Try to convert to numeric
            numeric_col = pd.to_numeric(df.iloc[:, col_idx], errors='coerce')
            # Check if any values couldn't be converted
            if numeric_col.isnull().any():
                non_numeric_issues.append(f"Column '{col_name}' (column {col_idx + 1}) contains non-numeric values")
        except Exception as e:
            non_numeric_issues.append(f"Column '{col_name}' (column {col_idx + 1}) has invalid data: {str(e)}")
    
    if non_numeric_issues:
        return None, "Non-numeric values found: " + "; ".join(non_numeric_issues)
    
    return df, None

def topsis(df, weights, impacts):
    """Calculate TOPSIS scores and ranks"""
    try:
        # Convert numeric columns to float
        data = df.iloc[:, 1:].astype(float)
        
        # Verify column count matches weights and impacts
        num_criteria = data.shape[1]
        if num_criteria != len(weights):
            return None, f"Number of criteria ({num_criteria}) does not match number of weights ({len(weights)})"
        if num_criteria != len(impacts):
            return None, f"Number of criteria ({num_criteria}) does not match number of impacts ({len(impacts)})"
        
        # Normalization
        norm = data / np.sqrt((data ** 2).sum())
        weighted = norm * weights

        # Ideal best and worst solutions
        ideal_best = []
        ideal_worst = []

        for i, imp in enumerate(impacts):
            if imp == '+':
                ideal_best.append(weighted.iloc[:, i].max())
                ideal_worst.append(weighted.iloc[:, i].min())
            else:
                ideal_best.append(weighted.iloc[:, i].min())
                ideal_worst.append(weighted.iloc[:, i].max())

        # Separation measures
        S_plus = np.sqrt(((weighted - ideal_best) ** 2).sum(axis=1))
        S_minus = np.sqrt(((weighted - ideal_worst) ** 2).sum(axis=1))

        # TOPSIS score
        score = S_minus / (S_plus + S_minus)

        df["Topsis Score"] = score
        df["Rank"] = score.rank(ascending=False, method="dense").astype(int)

        return df, None
    except Exception as e:
        return None, str(e)

def send_email(receiver, file_path, sender_email, app_password):
    """Send result file via email"""
    try:
        msg = EmailMessage()
        msg["Subject"] = "TOPSIS Analysis Result"
        msg["From"] = sender_email
        msg["To"] = receiver
        msg.set_content("Dear User,\n\nPlease find attached your TOPSIS analysis result file.\n\nBest regards,\nTOPSIS Web Service")

        with open(file_path, "rb") as f:
            msg.add_attachment(f.read(), maintype="application",
                               subtype="octet-stream",
                               filename="topsis_result.csv")

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
            server.login(sender_email, app_password)
            server.send_message(msg)
        return True, None
    except Exception as e:
        return False, str(e)

@app.route("/")
def index():
    return render_template("index.html")

@app.route("/submit", methods=["POST"])
def submit():
    """Handle TOPSIS calculation with email or webpage display"""
    try:
        # Check required parameters
        required_fields = ["file", "weights", "impacts", "result_option"]
        missing_fields = [field for field in required_fields if field not in request.form and field != "file"]
        
        if "file" not in request.files:
            missing_fields.append("file")
        
        if missing_fields:
            return jsonify({"status": "error", "message": f"Missing required parameters: {', '.join(missing_fields)}"}), 400

        # Get file and parameters
        file = request.files["file"]
        weights_str = request.form.get("weights", "").strip()
        impacts_str = request.form.get("impacts", "").strip()
        result_option = request.form.get("result_option", "email").strip()
        email = request.form.get("email", "").strip()

        # Validate file was selected
        if file.filename == "":
            return jsonify({"status": "error", "message": "No file selected. Please upload a CSV or Excel file."}), 400

        # Validate result option
        if result_option not in ["email", "display"]:
            return jsonify({"status": "error", "message": "Invalid result option. Choose 'email' or 'display'"}), 400

        # Validate email only if option is email
        if result_option == "email":
            if not email:
                return jsonify({"status": "error", "message": "Email ID is required for email delivery"}), 400
            if not validate_email(email):
                return jsonify({"status": "error", "message": f"Invalid email format: '{email}'. Please use format: user@example.com"}), 400

        # Validate weights
        weights, error = validate_weights(weights_str)
        if error:
            return jsonify({"status": "error", "message": f"Weights validation failed: {error}"}), 400

        # Validate impacts
        impacts, error = validate_impacts(impacts_str)
        if error:
            return jsonify({"status": "error", "message": f"Impacts validation failed: {error}"}), 400

        # Check if counts match
        if len(weights) != len(impacts):
            return jsonify({"status": "error", "message": f"Mismatch: {len(weights)} weights but {len(impacts)} impacts. They must be equal."}), 400

        # Read and validate file (CSV or Excel)
        try:
            if not file:
                return jsonify({"status": "error", "message": "File not found or cannot be read"}), 400
            
            df, error = read_file(file)
            if error:
                return jsonify({"status": "error", "message": f"File reading error: {error}"}), 400
            
        except FileNotFoundError:
            return jsonify({"status": "error", "message": "File not found. Please upload a valid CSV or Excel file."}), 400
        except pd.errors.EmptyDataError:
            return jsonify({"status": "error", "message": "File is empty. Please upload a file with data."}), 400
        except pd.errors.ParserError as e:
            return jsonify({"status": "error", "message": f"File parsing error: {str(e)}. Ensure file is properly formatted."}), 400
        except Exception as e:
            return jsonify({"status": "error", "message": f"Error reading file: {str(e)}"}), 400

        # Validate file structure
        df, error = validate_csv_file(df)
        if error:
            return jsonify({"status": "error", "message": f"File validation failed: {error}"}), 400

        # Check if number of criteria columns match weights and impacts
        num_criteria = df.shape[1] - 1
        if num_criteria != len(weights):
            return jsonify({"status": "error", "message": f"Mismatch: File has {num_criteria} criteria columns but {len(weights)} weights provided. They must match."}), 400
        if num_criteria != len(impacts):
            return jsonify({"status": "error", "message": f"Mismatch: File has {num_criteria} criteria columns but {len(impacts)} impacts provided. They must match."}), 400

        # Calculate TOPSIS
        result, error = topsis(df, weights, impacts)
        if error:
            return jsonify({"status": "error", "message": f"Error calculating TOPSIS: {error}"}), 500

        # Handle result based on option
        if result_option == "display":
            # Return result as JSON for webpage display
            result_data = result.to_dict('records')
            return jsonify({
                "status": "success",
                "message": "TOPSIS calculation completed successfully!",
                "result": result_data,
                "columns": list(result.columns)
            }), 200
        else:
            # Save result file and send via email
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            result_file = os.path.join(UPLOAD_FOLDER, f"result_{timestamp}.csv")
            result.to_csv(result_file, index=False)

            # Send email
            sender_email = os.getenv("SENDER_EMAIL", "")
            app_password = os.getenv("APP_PASSWORD", "")
            
            if not sender_email or not app_password:
                return jsonify({"status": "error", "message": "Email service not configured. Contact administrator."}), 500

            success, error = send_email(email, result_file, sender_email, app_password)
            if not success:
                return jsonify({"status": "error", "message": f"Error sending email: {error}"}), 500

            return jsonify({"status": "success", "message": "Result calculated and sent successfully to your email!"}), 200

    except Exception as e:
        return jsonify({"status": "error", "message": f"Unexpected error: {str(e)}"}), 500

if __name__ == "__main__":
    import os
    port = int(os.environ.get('PORT', 5000))
    debug = os.environ.get('FLASK_ENV') != 'production'
    app.run(host='0.0.0.0', port=port, debug=debug)
