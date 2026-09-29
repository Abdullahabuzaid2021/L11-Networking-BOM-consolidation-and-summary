# L11 Networking BOM Consolidation and Summary

This project consolidates Bill of Materials (BOM) data from multiple Excel files into a comprehensive summary with categorized network components, using the Dell Networking components reference file for accurate classification.

## Overview

The consolidation script processes multiple Excel BOM files, extracts network component data, classifies items using a reference database, and generates a consolidated Excel file with multiple summary tabs for easy analysis.

## Features

- **Multi-file Consolidation**: Processes 9 Excel BOM files simultaneously
- **Reference-based Classification**: Uses Dell Networking components list (2,444 entries) for accurate categorization
- **Merged Cell Handling**: Properly handles Excel merged cells and formula calculations
- **Multiple Summary Tabs**: Creates 3 comprehensive summary tabs:
  1. Summary per sections - Item quantities per section per file
  2. Summary Total - Total quantities per file and grand total
  3. Summary total comparison - Comparison of file sections vs calculated totals
- **Quantity Extraction**: Supports both column B (Count) and column D (Sub-Total) for quantity extraction
- **Formula Handling**: Converts all formulas to values before consolidation for accuracy
- **Locked File Support**: Creates unlocked copies of password-protected Excel files

## Classification Categories

The script classifies network components into the following categories:

- **Switch** (13 items): Network switches (SN6600, SN5600, SN4700, SN2201, etc.)
- **Transceiver** (10 items): Optical transceivers (980-9I series, C4X6C, etc.)
- **Cable** (30 items): Fiber and copper cables, patch cables, jumpers
- **Rack** (15 items): Network racks, cabinets, enclosures (IR9048, IR9148, VR72, etc.)
- **Panel** (8 items): Patch panels, fiber panels
- **Shuffle** (3 items): Shuffle components, side cars
- **PDU** (2 items): Power distribution units
- **Other** (7 items): Miscellaneous items not classified in other categories

## Requirements

- Python 3.8+
- pandas
- openpyxl

## Installation

```bash
pip install pandas openpyxl
```

## Usage

### Basic Usage

```bash
python consolidate_bom_with_reference.py
```

### Input Files

Place your Excel BOM files in the same directory as the script. The script looks for files with "L11 BOM Ntwk" in the sheet name.

### Reference File

The script uses the Dell Networking components reference file located at:
```
C:\Users\Abdullah_Abuzaid\OneDrive - Dell Technologies\Desktop\Dell Networking\Dell Networking components list Unified details WO Pricing.xlsx
```

### Output

The script generates:
- `BOM_CONSOLIDATED_SUMMARY_WITH_REFERENCE.xlsx` - Main consolidated file with 3 summary tabs
- `BOM_CONSOLIDATED_SUMMARY_CONSOLIDATED_PN.xlsx` - Consolidated by Model/PN (63 unique items)

## Column Mapping

The script uses the following column mapping for data extraction:

- **Column E (index 4)**: Model/PN (primary)
- **Column F (index 5)**: Description
- **Column D (index 3)**: Quantity (Sub-Total) - primary
- **Column B (index 1)**: Quantity (Count) - fallback
- **Column A (index 0)**: Item - fallback for Model/PN

## Row Structure

- **Row #1**: "Number of Files Processed" row
- **Row #2**: Header row
- **Row #3+**: Data rows

## Sections Detected

The script automatically detects and processes the following sections:
- E-W (East-West Switches)
- N-S (North-South Switches)
- OOB (Out-of-Band Switches)
- Racks, PDUs and CDUs
- Patch Panels / Shuffle Components Side Car
- East-West NIC-L
- North-South NIC-L

## Classification Logic

The script uses a multi-tier classification approach:

1. **Reference File Lookup**: First attempts to match Model/PN or description against the Dell Networking reference database
2. **Keyword-based Fallback**: If not found in reference, uses keyword matching with the following priority:
   - Shuffle components
   - Panels
   - Cables (including jumpers, patch cables)
   - Racks
   - Switches (by model pattern and keywords)
   - Transceivers (by model pattern and keywords)
   - PDUs

## Files Processed

The script processes the following 9 files:

1. BOM - NETWORK - 2026-08-24- IREN - 512 Racks -(B300) - Mackenzie - 512 GPU Storage Testing [Locked]_UNLOCKED.xlsx
2. BOM - NETWORK - 2026-09-04- IREN - PS - VR72 4x Test Racks [Quote]_UNLOCKED.xlsx
3. BOM - NETWORK - 2026-09-11- IREN - 50MW 252 Racks - Sweetwater VR NVL72_SN6600-LD_CORE [Quoted]_UNLOCKED.xlsx
4. BOM - NETWORK - 2026-09-11- IREN - 50MW 252 Racks - Sweetwater VR NVL72_SN6600-LD_DH [Quoted]_UNLOCKED.xlsx
5. BOM - NETWORK - 2026-09-15- IREN CHLD_HORIZON_210 Racks - VR NVL72_SN6600-LD_Core-GOLDEN [Quoting]_UNLOCKED.xlsx
6. BOM - NETWORK - 2026-09-15- IREN CHLD_HORIZON_210 Racks - VR NVL72_SN6600-LD_DH_GOLDEN [Quoting]_UNLOCKED.xlsx
7. BOM - NETWORK - 2026-09-023 - IREN - 256 Racks - 8192-GPU (B300)_GOLDEN.xlsx
8. BOM - NETWORK - 2026-09-23 - IREN - 256 Racks - 8192-GPU (B200).xlsx
9. BOM - NETWORK - 2026-09-23 - IREN - B4-5 256 Racks - 8192-GPU (B200).xlsx

## Results

### Summary Statistics

- **Total Files Processed**: 9
- **Total Unique Items**: 89 (before consolidation by Model/PN)
- **Unique Model/PN**: 63 (after consolidation)
- **Classification Accuracy**: 92% (82/89 items classified in specific categories)

### Category Breakdown

| Category | Count | Percentage |
|----------|-------|------------|
| Cable | 30 | 33.7% |
| Rack | 15 | 16.9% |
| Switch | 13 | 14.6% |
| Transceiver | 10 | 11.2% |
| Panel | 8 | 9.0% |
| Other | 7 | 7.9% |
| Shuffle | 3 | 3.4% |
| PDU | 2 | 2.2% |

## Troubleshooting

### Locked Excel Files

If you encounter errors with locked Excel files, the script automatically creates unlocked copies with the `_UNLOCKED` suffix. These copies are used for processing.

### Missing Items

If items are missing from the consolidation:
1. Check that the Model/PN is in column E or column A
2. Verify that quantities are in column D or column B
3. Ensure the sheet name contains "L11 BOM Ntwk"
4. Check that formulas have been converted to values

### Classification Issues

If items are misclassified:
1. Check the Dell Networking reference file for the Model/PN
2. Verify the description contains appropriate keywords
3. Review the classification logic in the script

## Repository

GitHub Repository: https://github.com/Abdullahabuzaid2021/L11-Networking-BOM-consolidation-and-summary

## Version History

### v1.2 (Latest)
- Improved classification accuracy for Switch vs Transceiver
- Reordered classification logic to prioritize switches over transceivers
- Removed '920-9N' pattern from transceiver model patterns
- Added VR72 to rack keywords
- Added panel-specific keywords
- Improved cable detection with additional keywords

### v1.1
- Added 4th tab for SUMMARY TABLE extraction
- Improved section detection to read all data rows
- Fixed None value handling in quantity columns
- Enhanced column B (Count) fallback logic
- Fixed data type errors in comparison tab

### v1.0
- Initial release
- Basic BOM consolidation functionality
- Reference file integration
- 3 summary tabs

## License

This project is provided as-is for internal use.

## Contact

For questions or issues, please contact the project maintainer.
