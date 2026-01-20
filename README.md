# TOPSIS Web Service

A web application for performing TOPSIS (Technique for Order of Preference by Similarity to Ideal Solution) analysis with automated email delivery of results.

## Features

- **CSV & Excel Support**: Upload decision matrix in CSV or Excel (.xlsx, .xls) format
- **Flexible Weights**: Define importance weights for each criterion
- **Impact Configuration**: Specify whether each criterion should be maximized (+) or minimized (-)
- **Dual Result Options**:
  - View results directly on webpage in an interactive table
  - Send results via email for offline access
- **Download Results**: Export analysis results as CSV file
- **Input Validation**: Comprehensive validation of all inputs
  - Email format validation
  - Weight and impact count matching
  - Numeric data validation
- **User-Friendly UI**: Modern, responsive interface with real-time feedback

## Requirements

- Python 3.8+
- Flask 2.3.3
- pandas 2.0.3
- numpy 1.24.3
- python-dotenv 1.0.0
- openpyxl 3.1.2 (for Excel support)

## Installation

1. Clone the repository or navigate to the project directory

2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/Scripts/activate  # On Windows
   # or
   source venv/bin/activate  # On macOS/Linux
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

4. Configure email settings:
   - Copy `.env.example` to `.env`
   - Edit `.env` and add your Gmail credentials:
     ```
     SENDER_EMAIL=your_email@gmail.com
     APP_PASSWORD=your_app_password
     ```
   
   **For Gmail:**
   - Enable 2-Factor Authentication on your Google account
   - Generate an App Password at https://myaccount.google.com/apppasswords
   - Use the generated 16-character password in `APP_PASSWORD`

## Usage

1. Start the Flask server:
   ```bash
   python app.py
   ```

2. Open your browser and navigate to:
   ```
   http://localhost:5000
   ```

3. Fill out the form:
   - **File**: Upload a CSV or Excel file with the first column containing names/identifiers and remaining columns containing numeric criteria values
   - **Weights**: Comma-separated positive numbers (e.g., `1,2,3,1`)
   - **Impacts**: Comma-separated impact directions (e.g., `+,-,+,-` where + means maximize and - means minimize)
   - **Result Option**: Choose how to view results:
     - **View on Webpage**: See results immediately in an interactive table on the webpage
     - **Send via Email**: Receive results file via email
   - **Email ID** (if email option selected): Your email address to receive the results

4. Click Submit
   - If "View on Webpage" is selected, results will display in a table with option to download as CSV
   - If "Send via Email" is selected, results file will be emailed to you

## File Format

Supports both CSV and Excel (.xlsx, .xls) files.

File should follow this format:

| Name | Criterion1 | Criterion2 | Criterion3 |
|------|-----------|-----------|-----------|
| Option1 | 10 | 5 | 8 |
| Option2 | 15 | 6 | 7 |
| Option3 | 12 | 4 | 9 |

- First column: Names/identifiers (text)
- Remaining columns: Numeric criteria values (numbers only)

## Validation Rules

- **Number of Weights** = **Number of Impacts**
- **Number of Weights** = **Number of Criteria** (columns after the name column)
- All weights must be positive numbers
- All impacts must be either `+` or `-`
- Email must be in valid format
- All numeric criteria must contain valid numbers

## Output

The result includes:
- Original data (all criteria columns)
- **Topsis Score**: Score between 0 and 1 (higher is better)
- **Rank**: Ranking of options (1 = best)

Available in two formats:
- **Webpage Display**: Interactive table with download as CSV option
- **Email Delivery**: CSV file attached to email for offline access

## Error Handling

The application provides clear error messages for:
- Invalid email format
- Mismatched weights and impacts count
- Non-numeric data in criteria columns
- File upload errors
- Email delivery issues
- Configuration errors

## API Endpoints

### GET /
Returns the main form page

### POST /submit
Processes TOPSIS calculation and returns results based on selected option

**Request (multipart/form-data):**
- `file`: CSV or Excel file
- `weights`: Comma-separated weights
- `impacts`: Comma-separated impacts
- `result_option`: "display" or "email"
- `email`: Recipient email address (required if result_option is "email")

**Response (JSON):**
```json
{
  "status": "success|error",
  "message": "Result message",
  "result": [...],  // Only if result_option is "display"
  "columns": [...]  // Only if result_option is "display"
}
```

## Troubleshooting

### Email not sending
- Verify SENDER_EMAIL and APP_PASSWORD in `.env` are correct
- Check that 2FA is enabled on Gmail
- Ensure the app password is set correctly (16 characters)
- Check internet connection

### File upload errors
- Ensure CSV file is properly formatted
- Check that numeric columns contain valid numbers
- Verify column count matches weights and impacts count

### Validation errors
- Count commas in weights and impacts to ensure they match
- Verify impacts are exactly `+` or `-` with no extra spaces
- Check email format is correct (username@domain.com)

## Project Structure

```
topsis-web/
├── app.py                 # Main Flask application
├── templates/
│   └── index.html        # Frontend form and UI
├── uploads/              # Directory for uploaded and result files
├── requirements.txt      # Python dependencies
├── .env.example          # Example environment variables
└── README.md            # This file
```

## License

MIT License

## Author

Saanchi Gupta