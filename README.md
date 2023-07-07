# Manufacturer DTC Analyzer

Manufacturer DTC Analyzer is a GUI application built with the `tkinter` library in Python. It allows users to process a summary table of fault codes and generates corresponding comments based on the analysis.

## Features

- Select and browse input files, including summary table, SAE J2012DA, and manufacturer DTC files.
- Configure the column indices for fault codes and the starting row.
- Specify the control unit for analysis.
- Process the summary table and generate comments based on fault code analysis.
- Save the processed output to an Excel file.

## Installation

1. Clone or download the repository:

   ```bash
   git clone https://github.com/zVeqsxy/DTC-Report-Analysis-Tool.git
   ```

2. Navigate to the `Manufacturer-DTC-Analyzer` directory:

   ```bash
   cd Manufacturer-DTC-Analyzer
   ```

3. Install the required dependencies by running the following command:

   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:

   ```bash
   python analyzer.py
   ```

2. The GUI window will appear, allowing you to interact with the application.

3. Follow the on-screen instructions to select input files, configure settings, and process the summary table.

4. The processed output will be saved as `Abgleich_herstellerspezifische_DTCs_{control_unit}.xlsx`.

## Contributing

Contributions to this project are welcome. If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).

## Acknowledgements

The Manufacturer DTC Analyzer application was developed using Python and the tkinter library.

## Contact

For any inquiries or questions, please contact [Ali Almaliki](mailto:Reyhamudi609@gmail.com).

