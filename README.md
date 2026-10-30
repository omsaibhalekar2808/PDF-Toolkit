# Python PDF Toolkit

A web-based PDF toolkit built with Python and Flask. The project brings common PDF processing operations together in one application.

## Features

- PDF processing and conversion tools supported by the application
- Web interface powered by Flask and HTML templates
- File upload handling and temporary-file processing
- Configurable upload limits
- Rate limiting and security-related application settings
- Production entry point through WSGI

> **Note:** The available tools depend on the functions implemented in `pdf_tools.py`. Update this list to match the features that have been tested.

## Project Structure

```text
github-project/
├── .venv/                  # Local virtual environment (not committed)
├── .env                    # Local secrets (not committed)
├── .env.example            # Example configuration
├── .gitignore
├── README.md
└── pdf_toolkit/
    ├── app.py              # Flask application and routes
    ├── pdf_tools.py        # PDF processing functions
    ├── requirements.txt    # Python dependencies
    ├── wsgi.py             # WSGI entry point
    └── templates/          # HTML templates
```

## Requirements

- Python 3.11
- pip
- The native libraries required by any PDF conversion dependencies, including WeasyPrint if enabled

## Installation

### 1. Clone the repository

```bash
git clone YOUR_GITHUB_REPOSITORY_URL
cd YOUR_REPOSITORY_FOLDER
```

Replace the placeholders with your actual repository URL and folder name.

### 2. Create a virtual environment

**Windows PowerShell**

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 3. Install dependencies

```powershell
python -m pip install --upgrade pip
python -m pip install -r .\pdf_toolkit\requirements.txt
```

### 4. Configure environment variables

Copy `.env.example` to `.env` and update the values for your local environment.

Do not commit `.env` or any real secret keys to GitHub.

### 5. Start the application

From the repository root:

```powershell
python .\pdf_toolkit\app.py
```

Open the local URL printed in the terminal.

## Configuration

Configure application settings using environment variables supported by the application.

For local development, use a `.env` file if the application is configured to load one. For production, set environment variables through your hosting provider.

Keep debug mode disabled on public deployments.

## WeasyPrint on Windows

WeasyPrint requires compatible native libraries in addition to its Python packages. Installing the Python package alone may not be sufficient.

Follow the official installation and troubleshooting instructions:

https://doc.courtbouillon.org/weasyprint/stable/first_steps.html

Do not rely on a machine-specific DLL path for deployment. Configure the native dependencies for the target operating system and hosting environment.

## Production Deployment

Use a production WSGI server supported by your hosting platform and point it to the application's WSGI entry point.

For example, if `wsgi.py` is inside `pdf_toolkit/` and exports `app`:

```text
pdf_toolkit.wsgi:app
```

Confirm the module path and exported variable match your actual code.

Before exposing the application publicly:

- Disable debug mode.
- Set a strong secret key.
- Configure request-size and rate limits.
- Validate uploaded files and restrict accepted file types.
- Store temporary files safely and clean them up.
- Restrict remote URL fetching when converting HTML to PDF.
- Test every conversion tool.
- Review dependencies for known vulnerabilities.
- Configure HTTPS and production logging.

## Security

This project is intended to process user-provided documents. Treat uploaded files as untrusted input.

Do not deploy publicly until file handling, access controls, resource limits, error handling, and PDF conversion dependencies have been tested.

## Troubleshooting

### Python imports fail in VS Code

Select the project's `.venv` interpreter and reinstall dependencies using the project's `requirements.txt`.

### WeasyPrint fails to load native libraries

Follow the official Windows installation guide linked above and verify that the required native dependencies are available.

### Dependency conflicts

Run:

```powershell
python -m pip check
```

## Contributing

Contributions and bug reports are welcome. Please include reproducible steps when reporting a problem. Avoid including private documents, secrets, or sensitive information in bug reports.

## License

Choose a license before distributing this project. If you intend to publish it as open source, add the appropriate license file to the repository.

## Disclaimer

This project is provided as-is. Test the application and its dependencies before using it with sensitive or important documents.
