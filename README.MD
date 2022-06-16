# Certificate Generator

## Dependencies

- [LibreOffice](https://www.libreoffice.org/)
- [Python 3.7+](https://www.python.org/)

## Usage

  1. Install dependencies
      - Linux - APT

          1. Install LibreOffice

              ```console
              sudo apt update
              sudo apt install libreoffice
              ```

          2. Create virtual environment and install dependencies

              ```console
              python3 -m venv venv
              source venv/bin/activate
              pip install -r requirements.txt
              ```

  2. Create a `names.txt` file with the names of the people you want to display;
  3. Create a `model.pptx` file with the certificate model slide;
     - The fields with `{{name}}` will be replaced with the person name;
  4. Run the program

      ```console
      python main.py ./output_file.pdf
      ```
