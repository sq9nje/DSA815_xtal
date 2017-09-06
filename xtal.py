"""XTAL parameter measurement automation using Rigol DSA815-TG.

"""

import sys, getopt
import math
import visa
from openpyxl.workbook import Workbook
from openpyxl.styles import Font

def print_banner():
    """Prints out a nice banner.
    
    """
    print('##################################')
    print('         Rigol DSA815-TG')
    print('    XTAL Parameter Measurement')
    print('##################################')
    print()

def get_idn(instrument):
    """Querries the instrument for its ID string using the *IDN? SCPI command.

    Args:
        param1: VISA instrument object
    Returns:
        string: Instrument's ID string
    
    """
    idn = instrument.query('*IDN?')
    return idn

def check_tg(instrument):
    """ Checks if tracking generator is enabled.

    Args:
        param1: VISA instrument object
    Returns:
        bool: True if TG is enabled

    """
    if int(instrument.query(':OUTPUT:STATE?').strip('\n\r')) == 0:
        return False
    else:
        return True

def setup_markers(instrument):
    """Enables trace markers and frequency display.
    Marker 1 features:
        - Track peak
        - -3dB bandwidth
    Frequency display:
        - 1Hz resolution

    Args:
        param1: VISA instrument object

    """
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

    # output file name
    ws_filename = ''
    # offset in xtal numbering
    start_offset = 0;

    # Process command line options
    opts, args = getopt.getopt(sys.argv[1:], "hf:", ["file="])
    for opt, arg in opts:
        if opt == '-h':
            print("Usage: xtal.py -f <file_name> -n <starting_number>\n")
            sys.exit(2)
        elif opt in ('-f', '--file'):
            ws_filename = arg
        elif opt in ('-n', '--number'):
            try:
                start_offset = int(arg) - 1
            except:
                print("Invalid starting number given! Using 1.")

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
    print('Device ID: {}\n'.format(get_idn(rigol)))
    
    # Check if tracking generator is enabled
    if not check_tg(rigol):
        print('!!! Tracking generator not enabled !!!')
        print('Please refer to the usage manual for instructions on how to setup the instrument')
        sys.exit(-1)
    
    # Configure marker functions
    setup_markers(rigol)

    # Create a new Excel workbook
    wb = Workbook()
    # New worksheet
    ws = wb.active
    ws.title = 'Measurements'

    # Column headers
    ws.append(('No.', 'Center Frequency', '-3dB Bandwidth', 'Attenuation', 'Q', 'Rm', 'Lm', 'Cm'))
    
    # Column headers in bold italic
    fnt = Font(bold = True, italic = True)
    for cell in ws["1:1"]:
        cell.font = fnt
    
    # Worksheet row counter
    #ws_row = 1

    xtal_number = start_offset

    # Measure next oscillator
    while raw_input('Measure XTAL? [Y/n]: ').lower() != 'n':
        xtal_number = xtal_number + 1
        fc = float(rigol.query('calc:marker:fcount:x?').strip('\n\r'))
        bw = float(rigol.query('calc:bandwidth:result?').strip('\n\r'))
        att = float(rigol.query('calc:marker1:y?').strip('\n\r'))
        
        rm = 25 * (math.pow(10, bw/20) - 1)
        reff = 25 + rm
        q = fc/bw
        lm = reff / (2 * math.pi * bw)
        cm = bw / 2 * math.pi * reff * fc**2

        # Append values to worksheet
        ws.append((xtal_number, fc, bw, att, q, rm, lm, cm))

        # Print values
        print('No.: %d Fc: %f   BW: %f   Att: %f   Q: %f   Rm: %f   Lm: %f   Cm: %f\n'%(xtal_number, fc, bw, att, q, rm, lm, cm))

    # Save the file
    if wb_filename == '':
        wb_filename = raw_input('Save as: ')
    wb.save(filename = wb_filename)
