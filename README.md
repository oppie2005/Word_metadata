# Word Metadata

Updates the metadata "Title" to match the word document filename. This is intended to be used during
the document finalization process, just prior to publishing the Word document to PDF format. 

Desired changes are to add a configuration file for changing other metadata items
that will be consistent across all intended documents or can be changed as needed. 

---

## Table of Contents

- [Word Metadata](#word-metadata)
  - [Table of Contents](#table-of-contents)
  - [Installation](#installation)
  - [Usage](#usage)
  - [Contributing](#contributing)
  - [License](#license)

---

## Installation

1. Make sure you have Python 3 installed.

```bash
python --version
```

2. Install Dependencies:

```bash
python -m pip install pywin32
```

---

## Usage

1. Drag and drop your Word file onto the word_metadata.py file. A command prompt will open and your file will be processed. 
   *You may see Word open and close during the process.*

**or**

1. Navigate to the folder which contains the word_metadata.py file.
2. Right-click anywhere in your folder and select "Open in terminal". 
3. Type the below into the command prompt, followed by the [ENTER] key.

   ```bash
   python word_metadata.py 'path to word file'
   ```

---

## Contributing

1. Fork the repository.
2. Create your feature branch:
   ```
   git checkout -b feature/YourFeature
   ```
3. Commit your changes:
   ```
   git commit -m 'Add some feature'
   ```
4. Push to the branch:
   ```
   git push origin feature/YourFeature
   ```
5. Open a pull request.

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

