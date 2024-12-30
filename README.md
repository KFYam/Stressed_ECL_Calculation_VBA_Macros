# Stressed_ECL_Calculation_VBA_Macros
The VBA macros are designed to manage and calculate Expected Credit Loss (ECL) under various stress scenarios, as part of financial risk management and compliance with accounting standards IFRS 9. The ECL represents potential losses from credit risk over the lifetime of financial assets (e.g., loans).

## Features
- Updates `FLAG_STAT_STAGE2` based on input conditions.
- Clears contents of specified ranges in `Stage2_STAT_StressedECL`.
- Provides dictionary functions to retrieve structured input data and mappings.
- Calculates partial expected life and adjusted metrics for different economic scenarios.

## Functions
  - UpdateStageSEQ: Updates stage sequences based on Stage1_STAT_StressedECL data.
  - Stage2Clear: Clears contents in Stage2_STAT_StressedECL.
  - DictInputField: Retrieves field mappings for Input_Data.
  - Gen_stat_Stage2_ECL: Main macro for generating stage 2 stressed ECL.**

## Explanations
1. Data Setup and Initialization
   Code Segments:
   DictInputField
   DictS2ECLField
   DictInputPD
   DictInputPWA
   DictLGD_nonretail
   DictLGD_retail
   Input fields: For example, FLAG_STAT_STAGE2, RATING_KEY, etc., to their column positions.
   Stress parameters: Like PDs, LGDs, and multipliers from the respective sheets (Input_PD, Input_PWA (probability weighted averages, etc.).

2. Stage 1 to Stage 2 Transition Identification
   UpdateStageSEQ
   Purpose:
     Analyze Stage1_STAT_StressedECL for indicators (SC1_SE1_Stage1_to_2n3_Ind, SC2_SE2_Stage1_to_2n3_Ind).
     Summarize conditions to determine whether an asset transitions to Stage 2.

3. Clearing Stage 2 Data
   Stage2Clear

   Purpose:
     Clear old data from the Stage2_STAT_StressedECL sheet before recalculations.

4. Static Data Copying
   Stage2_Static
   Purpose:
     Copy relevant static fields (e.g., Exposure Reference, SEQ, etc.) from Input_Data to Stage2_STAT_StressedECL based on FLAG_STAT_STAGE2.

5. Dynamic Calculations of Stress Scenarios
   Gen_stat_Stage2_ECL

   Purpose:
   Macro-Economic Scenarios:
      Iterate through combinations of scenarios (SC1, SC2, etc.) and severities (SE1, SE2, etc.).
  ECL Components:
      For each scenario-year combination, calculate:
        - Adjusted PD values (ST_PD_Good, etc.).
        - Stressed LGD values for retail and non-retail exposures.
        - PWA adjustments for weighting macroeconomic impacts.
        - Partial expected life for staged assets.
        - Aggregate ECL.
  Summation:
      Aggregate yearly ECL values into final scenario-based results (SC1_SE1_ECL, etc.).

## Summary of Key Steps
1. Input Data Preparation
    - Load staging and credit data from input sheets (Input_Data, Input_PD, etc.).
    - Create dictionaries for quick field and parameter lookups.
2. Stage Identification
    - Identify transitions from Stage 1 to Stage 2 using risk indicators.
    - Update staging flags dynamically.
3. Data Reset and Initialization  
    - Clear old Stage 2 results.
    - Copy relevant static fields to the target sheet.
4. Stress Scenario Calculations
    - Iterate through scenarios and severities to adjust risk metrics:
      PD: Probability of Default.
      LGD: Loss Given Default.
      EAD: Exposure at Default.
      PWA: Probability-Weighted Averages.
5. Calculate yearly and aggregate ECL values.

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
