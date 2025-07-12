# File To Docx

A simple Python script to generate a DOCX (Microsoft Word) file from multiple code files for your presentations or assignments.

![code-to-docx](https://socialify.git.ci/SharafatKarim/code-to-docx/image?description=1&descriptionEditable=A+docx+file+from+multiple+files+of+codes+for+your+assignment+or+presentation.&font=KoHo&forks=1&issues=1&language=1&name=1&owner=1&pattern=Plus&pulls=1&stargazers=1&theme=Auto)

## 🚀 Run Locally

1. **Clone the project**

  ```bash
  git clone https://github.com/SharafatKarim/code-to-docx
  ```

2. **Navigate to the project directory**

  ```bash
  cd code-to-docx
  ```

3. **Create a virtual environment**

  ```bash
  python -m venv .venv
  ```

4. **Activate the virtual environment**

- **Windows**

    ```bash
    .venv\Scripts\activate
    ```

- **Linux or macOS**

    ```bash
    source .venv/bin/activate
    ```

5. **Install dependencies**

  ```bash
  pip install -r requirements.txt
  ```

6. **Place your files in the `input` directory and run the script**

  ```bash
  python main.py
  ```

## 🖥️ GUI Version

A GUI version is available for easy input and output folder selection. Simply run:

```bash
python GUI.py
```

### ✨ New Features

- Add files to `blacklist.txt` to ignore them during processing.
- No need to specify file formats on the first run; the script will automatically search through all files and folders inside `input/`.

```bash
python main.py
```

## 🎥 Sample Video

[Watch a demo for CLI version](https://github.com/SharafatKarim/code-to-docx/assets/93897936/dcdd878a-cbc1-4c28-82c5-416a7a5146bd)

## 👥 Authors

[![SharafatKarim's Profilator](https://profilator.deno.dev/SharafatKarim?v=1.0.0.alpha.4)](https://github.com/SharafatKarim)
[![AyakashiKitsune's Profilator](https://profilator.deno.dev/AyakashiKitsune?v=1.0.0.alpha.4)](https://github.com/AyakashiKitsune)

## 🤝 Contribution

Feel free to fork the repository and open pull requests. Contributions, patches, and advanced versions are welcome—your imagination is the limit!
