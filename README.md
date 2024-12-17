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