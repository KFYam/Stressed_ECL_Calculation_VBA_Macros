# Stressed_ECL_Calculation_VBA_Macros
This repository contains VBA code to automate the computation and management of expected credit loss (ECL) calculations based on various input and staging criteria.

## Features
- Updates `FLAG_STAT_STAGE2` based on input conditions.
- Clears contents of specified ranges in `Stage2_STAT_StressedECL`.
- Provides dictionary functions to retrieve structured input data and mappings.
- Calculates partial expected life and adjusted metrics for different economic scenarios.

## Requirements
- Microsoft Excel with VBA support enabled.
- The following worksheets must exist in your workbook:
  - `Stage1_STAT_StressedECL`
  - `Stage2_STAT_StressedECL`
  - `Input_Data`
  - `Input_PD`
  - `Input_PWA`
  - `Input_stressed_LGD_multiplers`
  - `Input_Retail_stressedLGD`
 
## Installation
1. Clone the repository:
   ```bash
   git clone https://github.com/<username>/<repository-name>.git

2. Import the .bas file into your Excel workbook:
  - Open Excel.
  - Press ALT + F11 to open the VBA editor.
  - Right-click on "Modules" and select "Import File".
  - Choose the .bas file from the cloned repository.

## Usage
1. Load the macro-enabled workbook.
2. Run the macros in the VBA editor or via the Developer tab in Excel.

## Functions
  - UpdateStageSEQ: Updates stage sequences based on Stage1_STAT_StressedECL data.
  - Stage2Clear: Clears contents in Stage2_STAT_StressedECL.
  - DictInputField: Retrieves field mappings for Input_Data.
  - Gen_stat_Stage2_ECL: Main macro for generating stage 2 stressed ECL.**

