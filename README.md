# pdfreader

Simple command-line utility to extract text from a PDF file using `pypdf`.

## Setup

1. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

Run the script with a path to a PDF file:

```bash
python src/main.py <path-to-pdf>
```

An example using the included sample invoice:

```bash
python src/main.py samples/invoice.pdf
```

## Order comparison

Compare the latest order file against the generated invoice CSV and create
a color-coded Excel report:

```bash
python src/order_compare.py --order order/belso\ megrendeles.csv --invoice samples/invoice-output.csv --output order/compare-output.xlsx
```
