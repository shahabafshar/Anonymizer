# Excel Data Anonymizer

![photo_2023-09-14_07-19-37](https://github.com/shahabafshar/Anonymizer/assets/5617071/be77ac9f-629d-4cd0-9d90-33e9a42041f2)

### Description:
This utility is designed for data scientists and analysts who require draft datasets derived from original data without compromising privacy. With the Excel Data Anonymizer, users can generate a randomized version of their datasets, ensuring that the essential structures remain while the actual data values are altered. This utility is particularly handy when sharing datasets for public demos, examples, or initial stages of analysis where raw data can't be disclosed.

### Features:
- **Randomize Data**: Produces randomized values for numeric columns within a reasonable range of the original values, preserving the general structure and relationships in the data.
- **Preserve Formulas**: Cells containing formulas in the Excel sheet are kept intact, ensuring that the logic and calculations in the original data are not lost.
- **Intuitive GUI**: Offers an interactive and user-friendly interface, simplifying the anonymization process for users of all experience levels.

### Quick Start:
1. Clone the repository: `git clone https://github.com/shahabafshar/Anonymizer`
2. Navigate to the directory: `cd Anonymizer`
3. Install required libraries: `pip install -r requirements.txt`
4. Run the main file: `python excel_anonymizer.py`

### Setup:
Ensure you have Python installed. After cloning the repository, install the required libraries using the following command:
```
pip install -r requirements.txt
```

### Usage:
Run the main utility using the command:
```
python excel_anonymizer.py
```
Follow the GUI instructions to select an Excel file and anonymize its content.

### Contributions:
Feel free to submit pull requests, open issues, or provide feedback. We appreciate all contributions and feedback.

### License:
This project is licensed under the MIT License. See the `LICENSE` file for more details.

### Acknowledgments:
- Thanks to all contributors and the community for the valuable feedback and contributions.
