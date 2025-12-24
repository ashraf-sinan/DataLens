# Contributing to DataLens

Thank you for your interest in contributing to DataLens! This document provides guidelines for contributing to the project.

## How to Contribute

### Reporting Bugs

If you find a bug, please create an issue with:
- A clear, descriptive title
- Steps to reproduce the issue
- Expected behavior
- Actual behavior
- Screenshots (if applicable)
- Your environment (OS, Python version, etc.)

### Suggesting Features

Feature suggestions are welcome! Please create an issue with:
- A clear description of the feature
- Use case and benefits
- Any implementation ideas you may have

### Pull Requests

1. Fork the repository
2. Create a new branch (`git checkout -b feature/amazing-feature`)
3. Make your changes
4. Test your changes thoroughly
5. Commit your changes (`git commit -m 'Add amazing feature'`)
6. Push to the branch (`git push origin feature/amazing-feature`)
7. Open a Pull Request

### Development Setup

```bash
# Clone your fork
git clone https://github.com/yourusername/datalens.git
cd datalens

# Install dependencies
pip install -r requirements.txt

# Run the application
python src/main.py

# Run tests
python tests/test_export.py
```

### Code Style

- Follow PEP 8 style guidelines
- Use descriptive variable names
- Add comments for complex logic
- Write docstrings for functions and classes

### Testing

- Test your changes with various Excel files
- Test with single and multi-sheet workbooks
- Test with both quantitative and qualitative data
- Ensure the executable builds successfully

### Building

To build the Windows executable:

```bash
# Using the build script
build_app.bat

# Or manually
pyinstaller datalens.spec --clean
```

## Code of Conduct

- Be respectful and constructive
- Welcome newcomers and help them learn
- Focus on what is best for the community
- Show empathy towards other community members

## Questions?

Feel free to open an issue with your question, or reach out to the maintainers.

## License

By contributing to DataLens, you agree that your contributions will be licensed under the MIT License.
