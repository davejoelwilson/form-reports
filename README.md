# Form Reports

A Python tool to generate formatted reports from ConnectWise time entries.

## Features
- Fetches time entries from ConnectWise API
- Generates HTML debug reports
- Creates formatted Word documents
- Filters out meeting entries
- Handles weekly report generation (Monday-Friday)

## Requirements
- Python 3.8+
- ConnectWise API credentials
- Required Python packages (see requirements.txt)

## Setup
1. Install requirements:
    ```bash
    pip install -r requirements.txt
    ```

2. Configure environment variables:
    ```env
    CONNECTWISE_API_URL=your_api_url
    CONNECTWISE_COMPANY_ID=your_company_id
    CONNECTWISE_PUBLIC_KEY=your_public_key
    CONNECTWISE_PRIVATE_KEY=your_private_key
    ```

3. Configure settings in `config/settings.py`:
    ```python
    IGNITE_COMPANY_ID = "your_company_id"
    OUTPUT_DIR = "path/to/output"
    ```

## Usage
Run the main script:
bash
python connectwise_report/main.py


The script will:
1. Calculate last week's date range (Monday-Friday)
2. Fetch time entries from ConnectWise
3. Filter out meeting entries
4. Generate an HTML debug report
5. Create a formatted Word document

## Error Handling
- Handles API connection errors
- Validates date ranges
- Checks for empty time entries
- Creates output directory if not exists

## Contributing
1. Fork the repository
2. Create a feature branch
3. Submit a pull request