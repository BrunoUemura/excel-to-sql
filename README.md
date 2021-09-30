# Excel Sheet to SQL Insert Query

Convert Excel sheet to SQL Insert Query

## Setup

Create virtual environment:

```bash
virtualenv venv
source venv/Scripts/activate
```

Install the dependencies

```bash
pip install -r requirements.txt
```

Add a `.xlsx` file in `input` directory with the following structure: \
row 1: Table name \
row 2: Columns name \
row >= 3: Data to insert \
![alt text](./docs/sheet_example.png)

## How it works

After the `Setup` process is done, run the `main.py` script.

```bash
python main.py
```

The script will generate a `.txt` file in `output` directory with the SQL Insert query according to `.xlsx` file added in `input` directory.

## Author

- Bruno Hideki Uemura [linkedin](https://www.linkedin.com/in/bruno-hideki-uemura-918589139/), [Github](https://github.com/BrunoUemura)
