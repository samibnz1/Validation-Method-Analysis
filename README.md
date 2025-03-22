# List of Packages

- **pandas**: For data manipulation and analysis using DataFrames.
- **numpy**: For numerical calculations and matrix operations.
- **statsmodels**: For statistical analysis, linear regression, and hypothesis testing.
- **scipy**: For advanced statistical tests and other scientific computations.
- **matplotlib**: For data visualization in the form of graphs.
- **python-docx**: For creating reports in Word format.
- **pillow**: For image manipulation (useful for graphs).
- **tk**: For creating a simple graphical user interface.
- **openpyxl**: For reading and writing Excel files (used for importing data).

---

# 2. Excel Files and Data Structure

The project requires several Excel files to validate the analysis method according to different criteria. Here’s a guide to properly structure your Excel files:

### 1. Linearity Validation

- **Table for Active Ingredient (PA) only**: Contains data to validate linearity with the active ingredient alone.
- **Table for PA + FPR**: Contains data to validate linearity with the active ingredient and FPR.

#### Expected Structure for Each Linearity File:
| Factor   | x (independent variable) | y (response) |
|----------|--------------------------|--------------|
| Level 1  | 10                       | 12           |
| Level 2  | 20                       | 25           |
| Level 3  | 30                       | 35           |

### 2. Accuracy Validation

The Excel file for accuracy contains the data for PA + FPR.

#### Expected Structure for the Accuracy File:
| Level | xij | yij  | Reference |
|-------|-----|------|-----------|
| 1     | 10  | 9.8  | 10        |
| 2     | 20  | 19.5 | 20        |
| 3     | 30  | 29.2 | 30        |

### 3. Repeatability Validation

The Excel file for repeatability contains data for variability and repeatability between series.

#### Expected Structure for the Repeatability File:
| Series | Trial | Quantity Introduced | Air | Response |
|--------|-------|---------------------|-----|----------|
| 1      | 1     | 5                   | 2.1 | 5.2      |
| 1      | 2     | 5                   | 2.3 | 5.3      |
| 2      | 1     | 10                  | 3.0 | 10.1     |

> **Important:**  
> For the accuracy and repeatability validation, the "Day Constants" represent a calibration factor for each testing day. These constants are obtained by dividing the measured response (yij) by the introduced quantity (xij) for each repetition on the respective day.

- **Day Constant 1**: Calculated as the ratio of yij (measured response) to xij (introduced quantity) for day 1.
- **Day Constant 2**: Similarly calculated for day 2.
- **Day Constant 3**: Similarly calculated for day 3.

These constants are used to validate the accuracy of the results by adjusting the measurements based on the specific experimental conditions for each testing day.

### 4. Robustness Validation

The files for robustness contain a Hadamard matrix with the responses in the last column.

#### Expected Structure for the Robustness File:
| x1 | x2 | Response |
|----|----|----------|
| 1  | 0  | 10       |
| 0  | 1  | 12       |
| 1  | 1  | 11       |

### 5. Exactitude Profile

For exactness, you need to choose the number of levels and then import tables based on the number of levels specified. For each level, you will need to enter a reference value.

#### Expected Structure for the Exactness Files:
| Level | xij | yij  | Reference |
|-------|-----|------|-----------|
| 1     | 10  | 9.8  | 10        |
| 2     | 20  | 19.5 | 20        |
| 3     | 30  | 29.2 | 30        |

---

# 3. Usage Instructions

1. **Check each Excel file**: Ensure your Excel files are structured as per the examples provided above.
2. **Run the script**: Once the files are prepared, you can import these files into the script to perform the different validations.
3. **Analysis of Results**: The script will generate reports detailing the results of the various analyses (linearity, accuracy, repeatability, robustness, and exactness).

---

# Contact Information

If you encounter any problems or have questions, feel free to contact me:

- **Name**: Benzaimia Mohamed Sami  
- **Email**: [medsamibenzaimia@gmail.com](mailto:medsamibenzaimia@gmail.com)
