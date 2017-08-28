# XTAL parameter measurement with Rigol DSA815-TG
This is a tool that automates the process of measuring quarz crystal oscillator parameters using a Rigol DSA815-TG spectrum analyzer with the tracking generator option. You will also need a measuring fixture as described by Jack K8ZOA in "Crystal Motional Parameters: A Comparison of Measurement Approaches"

This tool uses what is known as the "-3dB method" to measure the crystal motional parameters. It lets you easily measure parameters of large batches of XTALs and produces a nice Excel spreadsheet.

### Prerequisites
- Rigol DSA815-TG spectrum analyzer with tracking gen
- Advanced Measurement Kit active license (required for NdB markers)
- Rigol SigmaView for SCPI controll

### Measurement fixture
A simple measurement fixture is required to measure the XTAL parameters using this method. This fixture consists of two PI attenuators that aim to match 50 ohm input and output of the spectrum analyzer to 12.5 ohms and a socket for the DUT. A short link is inserted in place of the crystal for calibration.

![Measurement Fixture Schematic](https://github.com/sq9nje/DSA815_xtal/raw/master/Fixture/images/XTAL_Set.sch.svg "Measurement Fixture Schematic")

Full KiCAD project documentation for this fixture is included in this repository. The board uses SMA connectors and 0805 1% resistors. A short piece of a precision IC socket can be used for the XTAL.

### Measurement procedure
1. Connect spectrum analyzer to computer using USB
2. Connect XTAL fixture to the spectrum analyzer
3. Enable tracking generator
4. Insert a link in place of an XTAL and perform TG normalisation
5. Insert sample XTAL into fixture and set frequency span and amplitude so that the resonance is clearly visible
6. Run script