# Validation-Method-Analysis

Analysis Method Validation - README

This project is a Python script designed to validate an analysis method using different statistical criteria. It performs tests on the provided data and generates detailed reports in Word format. 

The validation criteria include:

Linearity: Validating the linear relationship between variables.
Accuracy: Validating the method's precision against reference values.
Repeatability: Validating the consistency and variability of measurements.
Robustness: Validating the stability of the method using a Hadamard matrix.
Exactness: Validating the method according to specified levels and tolerance limits.
The script uses several Python packages to perform statistical analyses and generate reports with graphs.

1. Required Packages
Before running this project, you need to install the following packages. You can install them easily using pip by running the following command in your terminal or command prompt:


pip install pandas numpy statsmodels scipy matplotlib python-docx pillow tk openpyxl


List of Packages:
pandas: For data manipulation and analysis using DataFrames.
numpy: For numerical calculations and matrix operations.
statsmodels: For statistical analysis, linear regression, and hypothesis testing.
scipy: For advanced statistical tests and other scientific computations.
matplotlib: For data visualization in the form of graphs.
python-docx: For creating reports in Word format.
pillow: For image manipulation (useful for graphs).
tk: For creating a simple graphical user interface.
openpyxl: For reading and writing Excel files (used for importing data).


2. Excel Files and Data Structure
The project requires several Excel files to validate the analysis method according to different criteria. Here’s a guide to properly structure your Excel files:

1. Linearity Validation
Table for Active Ingredient (PA) only: Contains data to validate linearity with the active ingredient alone.
Table for PA + FPR: Contains data to validate linearity with the active ingredient and FPR.
Expected Structure for Each Linearity File:
Columns:
Factor or Level: The tested factor or level.
Response: The obtained measurement or response for this factor/level.
Independent Variable (x): The values of the independent factors.
Dependent Variable (y): The values of the response or the observed criterion.
Example:

Factor	x (independent variable)	y (response)
Level 1	10	12
Level 2	20	25
Level 3	30	35


2. Accuracy Validation
The Excel file for accuracy contains the data for PA + FPR.

Expected Structure for the Accuracy File:
Columns:
Level: The validation level.
xij: The measured factor.
yij: The response for this level.
Reference: The reference value for the validation level.
Example:

Level	xij	yij	Reference
1	10	9.8	10
2	20	19.5	20
3	30	29.2	30


3. Repeatability Validation
The Excel file for repeatability contains data for variability and repeatability between series.

Expected Structure for the Repeatability File:
Columns:
Series: Indicates the test series (e.g., Series 1, Series 2).
Trial: The identifier for the trial in the series.
Quantity Introduced: The quantity introduced in the test.
Air: The measurement of air or other factors.
Response: The measured response or result for this trial.
Example:

Series	Trial	Quantity Introduced	Air	Response
1	1	5	2.1	5.2
1	2	5	2.3	5.3
2	1	10	3.0	10.1


Important : 
For the accuracy and repeatability validation, the "Day Constants" represent a calibration factor for each testing day. These constants are obtained by dividing the measured response (yij) by the introduced quantity (xij) for each repetition on the respective day.

Day Constant 1: Calculated as the ratio of yij (measured response) to xij (introduced quantity) for day 1.
Day Constant 2: Similarly calculated for day 2.
Day Constant 3: Similarly calculated for day 3.
These constants are used to validate the accuracy of the results by adjusting the measurements based on the specific experimental conditions for each testing day.




4. Robustness Validation
The files for robustness contain a Hadamard matrix with the responses in the last column.

Expected Structure for the Robustness File:
Columns:
Factors: The different factors of the Hadamard matrix (e.g., x1, x2, etc.).
Response: The last column containing the responses or results of tests performed with the different factors.
Example:

x1	x2	Response
1	0	10
0	1	12
1	1	11


5. Exactitude Profil
For exactness, you need to choose the number of levels and then import tables based on the number of levels specified. For each level, you will need to enter a reference value.

Expected Structure for the Exactness Files:
Columns:
Level: The different validation levels.
xij: Measured value at this level.
yij: The response or measure of the variable of interest.
Reference: The reference value for this specific level.
Example (if you have 3 levels):

Level	xij	yij	Reference
1	10	9.8	10
2	20	19.5	20
3	30	29.2	30


3. Usage Instructions
Check each Excel file: Ensure your Excel files are structured as per the examples provided above.
Run the script: Once the files are prepared, you can import these files into the script to perform the different validations.
Analysis of Results: The script will generate reports detailing the results of the various analyses (linearity, accuracy, repeatability, robustness, and exactness).


Contact Information
If you encounter any problems or have questions, feel free to contact me:

Benzaimia Mohamed Sami
Email: medsamibenzaimia@gmail.com

