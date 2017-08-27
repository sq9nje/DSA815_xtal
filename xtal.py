import sys
import visa
from openpyxl.workbook import Workbook
from openpyxl.styles import Font

def print_banner():
    print('##################################')
    print('         Rigol DSA815-TG')
    print('    XTAL Parameter Measurement')
    print('##################################')
    print()

def print_idn(instrument):
    idn = instrument.query('*IDN?')
    print('Device ID: {}\n'.format(idn))

def check_tg(instrument):
    # Check if tracking generator is enabled
    if int(instrument.query(':OUTPUT:STATE?').strip('\n\r')) == 0:
        print('!!! Tracking generator not enabled !!!')
        print('Please refer to the usage manual for instructions on how to setup the instrument')
        sys.exit(-1)

def setup_markers(instrument):
    # Enable marker 1
    instrument.write('calc:marker1:state 1')
    # Marker 1 track peak
    instrument.write('calc:marker1:cpeak 1')
    # Marker 1 -3dB bandwidth measurement
    instrument.write('calc:marker1:function NDB')
    instrument.write('calc:bandwidth:ndb -3')
    # Enable frequency display
    instrument.write('calc:marker:fcount:state 1')
    # Set frequency resolution to 1Hz
    instrument.write('calc:marker:fcount:resolution:auto 0')
    instrument.write('calc:marker:fcount:resolution 1')



if __name__ == '__main__':
    print_banner()
    # Connect to first Rigol DSA device using VISA
    rm = visa.ResourceManager()
    dsa = list(filter(lambda x: 'DSA' in x, rm.list_resources()))
    if len(dsa) != 1:
        print('Bad instrument list: Could not find a Rigol DSA.')
        sys.exit(-1)

    print('Connecting to device...')
    rigol = rm.open_resource(dsa[0])
    rigol.timeout = 10000
    # Print device ID
    print_idn(rigol)
    # Check if tracking generator is enabled
    check_tg(rigol)
    # Configure marker functions
    setup_markers(rigol)

    # Create a new Excel workbook
    wb = Workbook()
    # New worksheet
    ws = wb.active
    ws.title = 'Measurements'

    # Column headers
    ws['A1'] = 'Center Frequency'
    ws['B1'] = '-3dB Bandwidth'
    ws['C1'] = 'Attenuation'

    # Column headers in bold italic
    fnt = Font(bold = True, italic = True)
    ws['A1'].font = fnt
    ws['B1'].font = fnt
    ws['C1'].font = fnt

    # Worksheet row counter
    ws_row = 1

    # Measure next oscillator
    while raw_input('Measure XTAL? [Y/n]: ') != 'n':
        ws_row = ws_row + 1

        fc = int(rigol.query('calc:marker:fcount:x?').strip('\n\r'))
        bw = int(rigol.query('calc:bandwidth:result?').strip('\n\r'))
        att = float(rigol.query('calc:marker1:y?').strip('\n\r'))

        # Append values to worksheet
        ws.cell(column = 1, row = ws_row, value = fc)
        ws.cell(column = 2, row = ws_row, value = bw)
        ws.cell(column = 3, row = ws_row, value = att)
        # Print values
        print('Fcenter: %d   BW: %d   Att: %f\n'%(fc,bw,att))

    wb_filename = raw_input('Save as: ')
    wb.save(filename = wb_filename)
