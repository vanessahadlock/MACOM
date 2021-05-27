## Conversion measurements for the MAMX-011054-EVAL.
## Eval Board Measurements

import xlsxwriter
import datetime
import time
import sys
import math
import numpy

from hw_qa_tools.visa_analyzer import SignalAnalyzer
from hw_qa_tools.visa_generator import SignalGenerator


##########################################################################################################
# Initialize the spreadsheet
# Returns the workbook object and spreadsheet name
# params:   spreadsheet_name_str
# returns:  workbook, spreadsheet_name
##########################################################################################################
def spreadsheet_setup(spreadsheet_name_str):

    # Create the spreadsheet
    current_time = datetime.datetime.now()
    spreadsheet_name = str(spreadsheet_name_str) + '_' + current_time.strftime("%Y-%m-%d_%H-%M") + '.xlsx'
    workbook = xlsxwriter.Workbook(spreadsheet_name)

    # Return the spreadsheet object
    return workbook, spreadsheet_name


##########################################################################################################
# Creates a notes page on the spreadsheet to track the spreadsheet test info and returns the notes sheet
# params:   workbook, test_notes, spreadsheet_name
# returns:  worksheet_notes
##########################################################################################################
def spreadsheet_test_info(workbook, test_notes, spreadsheet_name):

    # Create the notes sheet
    worksheet_notes = workbook.add_worksheet('Notes')
    # Write in the Notes
    row = 0
    worksheet_notes.write(row, 0, 'Date Time')
    current_time = datetime.datetime.now()
    worksheet_notes.write(row, 1, current_time.strftime("%Y-%m-%d %H:%M"))
    row = row + 1

    worksheet_notes.write(row, 0, 'Test Notes')
    worksheet_notes.write(row, 1, test_notes)
    row = row + 1

    worksheet_notes.write(row, 0, 'Spreadsheet Name')
    worksheet_notes.write(row, 1, spreadsheet_name)
    row = row + 1

    return worksheet_notes


##########################################################################################################
# Sweeps the mixer to test upconversion gain
# params:   workbook, if_freq, lo_freq, rf_freq, if_pin, lo_pin, if_mxg, lo_mxg, specan, if_loss, lo_loss, rf_loss
# returns:  worksheet_upconversion
##########################################################################################################
def upconversion_sweep(workbook, if_freq, lo_freq, rf_freq, if_pin, lo_pin, if_mxg, lo_mxg, specan, if_loss, lo_loss, rf_loss):

    # Empty list for the measurements
    if_frequencies = []
    lo_frequencies = []
    rf_frequencies = []
    lo_power = []
    raw_pout = []
    rf_losses = []

    # Turn on the mxgs
    if_mxg.on()
    lo_mxg.on()

    for i in range(0, len(if_freq)):
        # Set the IF mxg
        if_mxg.set_frequency(if_freq[i] * 1e9)
        time.sleep(0.5)
        # Set the mxg power levels
        if_mxg.set_amplitude(if_pin + if_loss[i])
        time.sleep(0.5)

        for k in range(0, len(lo_pin)):
            # Set the LO mxg
            lo_mxg.set_amplitude(lo_pin[k] + lo_loss)
            time.sleep(0.5)

            for j in range(0, len(rf_freq)):
                # Set the specan frequency
                specan.set_frequency(rf_freq[j] * 1e9)
                time.sleep(0.5)

                # Set the lo frequency
                lo_freq = (rf_freq[j] + if_freq[i])
                lo_mxg.set_frequency(lo_freq * 1e9)
                time.sleep(0.5)

                # Set the marker
                specan.set_marker(1, rf_freq[j] * 1e9)
                time.sleep(0.5)

                raw_pout.append(specan.get_power(1))
                if_frequencies.append(if_freq[i])
                lo_frequencies.append(lo_freq)
                rf_frequencies.append(rf_freq[j])
                lo_power.append(lo_pin[k])
                rf_losses.append(rf_loss[j])

    # Add a new page to the workbook
    worksheet_upconversion = workbook.add_worksheet('Up_Conversion')

    # Write the header information
    row = 0
    worksheet_upconversion.write(row, 0, 'IF Frequency (GHz)')
    worksheet_upconversion.write(row, 1, 'LO Frequency (GHz)')
    worksheet_upconversion.write(row, 2, 'RF Frequency (GHz)')
    worksheet_upconversion.write(row, 3, 'Specan Raw Pin (dBm)')
    worksheet_upconversion.write(row, 4, 'Eval Board RF Pout (dBm)')
    worksheet_upconversion.write(row, 5, 'Eval Board IF Pin (dBm)')
    worksheet_upconversion.write(row, 6, 'Eval Board LO Pin (dBm)')
    worksheet_upconversion.write(row, 7, 'Eval Board Conversion Gain (dB)')

    # Dump the data into the worksheet
    for i in range(0, len(raw_pout)):
        row = row + 1
        worksheet_upconversion.write(row, 0, if_frequencies[i])
        worksheet_upconversion.write(row, 1, lo_frequencies[i])
        worksheet_upconversion.write(row, 2, rf_frequencies[i])
        worksheet_upconversion.write(row, 3, raw_pout[i])
        worksheet_upconversion.write(row, 4, raw_pout[i] + rf_losses[i])
        worksheet_upconversion.write(row, 5, if_pin)
        worksheet_upconversion.write(row, 6, lo_power[i])
        worksheet_upconversion.write(row, 7, raw_pout[i] + rf_losses[i] - if_pin)

    # Return the sheet
    return worksheet_upconversion


##########################################################################################################
# Sweep the mixer to test downconversion gain
# params:   workbook, if_freq, lo_freq, rf_freq, rf_pin, lo_pin, rf_mxg, lo_mxg, specan, if_loss, lo_loss, rf_loss
# returns:  worksheet_downconversion
##########################################################################################################
def downconversion_sweep(workbook, if_freq, lo_freq, rf_freq, rf_pin, lo_pin, rf_mxg, lo_mxg, specan, if_loss, lo_loss, rf_loss):

    # Empty list for the measurements
    if_frequencies = []
    lo_frequencies = []
    rf_frequencies = []
    raw_pout = []
    lo_power = []
    raw_pout = []
    rf_losses = []

    # Set the mxg power levels
    rf_mxg.set_amplitude(rf_pin + rf_loss)
    lo_mxg.set_amplitude(lo_pin + lo_loss)

    # Turn on the mxgs
    rf_mxg.on()
    lo_mxg.on()

    for i in range(0, len(if_freq)):
        # Set the specan to the IF Frequency
        specan.set_frequency(if_freq[i] * 1e9)
        time.sleep(0.5)

        for j in range(0, len(rf_freq)):
            # Set the rf mxg
            rf_mxg.set_frequency(rf_freq[j] * 1e9)
            time.sleep(0.5)

            # Set the lo frequency
            lo_freq = (rf_freq[j] + if_freq[i]) / 4
            lo_mxg.set_frequency(lo_freq * 1e9)
            time.sleep(0.5)

            # Set the marker
            specan.set_marker(1, if_freq[i] * 1e9)
            time.sleep(0.5)

            raw_pout.append(specan.get_power(1))
            if_frequencies.append(if_freq[i])
            lo_frequencies.append(lo_freq)
            rf_frequencies.append(rf_freq[j])

    # Add a new page to the workbook
    worksheet_downconversion = workbook.add_worksheet('Down_Conversion')

    # Write the header information
    row = 0
    worksheet_downconversion.write(row, 0, 'IF Frequency (GHz)')
    worksheet_downconversion.write(row, 1, 'LO Frequency (GHz)')
    worksheet_downconversion.write(row, 2, 'RF Frequency (GHz)')
    worksheet_downconversion.write(row, 3, 'Specan Raw Pin (dBm)')
    worksheet_downconversion.write(row, 4, 'Eval Board IF Pout (dBm)')
    worksheet_downconversion.write(row, 5, 'Eval Board RF Pin (dBm)')
    worksheet_downconversion.write(row, 6, 'Eval Board LO Pin (dBm)')
    worksheet_downconversion.write(row, 7, 'Eval Board Conversion Gain (dB)')

    # Dump the data into the worksheet
    for i in range(0, len(raw_pout)):
        row = row + 1
        worksheet_downconversion.write(row, 0, if_frequencies[i])
        worksheet_downconversion.write(row, 1, lo_frequencies[i])
        worksheet_downconversion.write(row, 2, rf_frequencies[i])
        worksheet_downconversion.write(row, 3, raw_pout[i])
        worksheet_downconversion.write(row, 4, raw_pout[i] + if_loss)
        worksheet_downconversion.write(row, 5, rf_pin)
        worksheet_downconversion.write(row, 6, lo_pin)
        worksheet_downconversion.write(row, 7, raw_pout[i] + if_loss - rf_pin)

    # Return the sheet
    return worksheet_downconversion


##########################################################################################################
# params:   if_freq, rf_freq, sideband, mult
# returns:  lo_freq
##########################################################################################################
def synth_freq_gen(if_freq, rf_freq, sideband):
    """
    if_freq = list of IF frequencies
    rf_freq = list of RF frequencies
    sideband = upper or lower
    mult = LO multiplication
    """

    # Make the list for the lo frequencies
    lo_freq = []

    for i in range(0, len(if_freq)):
        for j in range(0, len(rf_freq)):
            if sideband == "upper":
                calculated_freq = (rf_freq[j] - if_freq[i])
                lo_freq.append(calculated_freq)
            if sideband == "lower":
                calculated_freq = (rf_freq[j] + if_freq[i])
                lo_freq.append(calculated_freq)
            else:
                return "Invalid Sideband"

    # Return the list
    return lo_freq


##########################################################################################################
# Function creates the list of powers to be swept over and returns a list
# params:   start_dbm, stop_dbm, step, max_pin_dbm
# returns:  power_sweep_values
##########################################################################################################
def power_sweep_range(start_dbm, stop_dbm, step, max_pin_dbm):

    # Empty list to store range
    power_sweep_values = []

    # Range between high and low powers
    range_db = stop_dbm - start_dbm

    current_value = start_dbm
    for i in range(0, math.ceil(range_db)):
        if current_value < max_pin_dbm:
            # Append the value to the loop
            power_sweep_values.append(round(current_value, 2))

            # Iterate the current value
            current_value = current_value + 1

        else:
            # If the max power is exceeded break out of the loop
            break

    # Return the range
    return power_sweep_values


##########################################################################################################
# Sweep the IF input power to see what the P1dB is
# params:   workbook, if_freq, lo_freq, rf_freq, if_pin, lo_pin, if_mxg, lo_mxg, specan, if_loss, lo_loss, rf_loss
# returns:  worksheet_upc_p1db, worksheet_upc_p1db_raw
##########################################################################################################
def tx_p1db(workbook, if_freq, lo_freq, rf_freq, if_pin, lo_pin, if_mxg, lo_mxg, specan, if_loss, lo_loss, rf_loss):

    # Set the mxg power levels
    if_mxg.set_amplitude(if_pin[0] + if_loss)
    lo_mxg.set_amplitude(lo_pin + lo_loss)

    # Turn on the mxgs
    if_mxg.on()
    lo_mxg.on()

    # Create the spreadsheet page
    worksheet_upc_p1db = workbook.add_worksheet('Up_Conversion_OP1dB')
    worksheet_upc_p1db_raw = workbook.add_worksheet('Up_Conversion_OP1dB_Raw')

    # Write the header information
    # One sheet is for the raw information and one is for the calibrated info
    row = 0
    worksheet_upc_p1db.write(row, 0, 'IF Pin (dBm)')

    worksheet_upc_p1db_raw.write(row, 0, 'IF Pin (dBm)')

    # Starting Column
    col = 0  # Starting at 0 instead of 1 because it immediately increments

    # Test loop
    for i in range(0, len(if_freq)):
        # Set the IF mxg
        if_mxg.set_frequency(if_freq[i] * 1e9)
        time.sleep(0.5)

        for j in range(0, len(rf_freq)):
            # Create / Blank the lists that store the data
            if_frequencies = []
            lo_frequencies = []
            rf_frequencies = []
            if_powers = []
            raw_pout = []

            # Reset the row for the current frequency
            row = 0

            # Increment to the next column
            col = col + 1

            # Set the specan frequency
            specan.set_frequency(rf_freq[j] * 1e9)
            time.sleep(0.5)

            # Set the lo frequency
            lo_freq = (rf_freq[j] + if_freq[i]) / 4
            lo_mxg.set_frequency(lo_freq * 1e9)
            time.sleep(0.5)

            # Set the marker
            specan.set_marker(1, rf_freq[j] * 1e9)
            time.sleep(0.5)

            # Write the column header information
            header_info = "Pout(dBm), RF {}GHz, LO {}GHz, IF {}GHz".format(rf_freq[j], lo_freq, if_freq[i])
            worksheet_upc_p1db.write(0, col, header_info)

            header_info_raw = "Pout_raw(dBm), RF {}GHz, LO {}GHz, IF {}GHz".format(rf_freq[j], lo_freq, if_freq[i])
            worksheet_upc_p1db_raw.write(0, col, header_info_raw)

            for k in range(0, len(if_pin)):
                # Sweep the IF power and set record the output power.
                # Set the IF power
                if_mxg.set_amplitude(if_pin[k] + if_loss)
                time.sleep(1)

                # Measure the output power
                raw_pout.append(specan.get_power(1))
                if_frequencies.append(if_freq[i])
                lo_frequencies.append(lo_freq)
                rf_frequencies.append(rf_freq[j])
                if_powers.append(if_pin[k])

                # Print the current Settings
                debug = "IF {}dBm, RF {}GHz, LO {}GHz, IF {}GHz".format(if_pin[k], rf_freq[j], lo_freq, if_freq[i])
                print(debug)

            # Dump the data into the spreadsheet
            for l in range(0, len(raw_pout)):
                row = row + 1
                worksheet_upc_p1db.write(row, 0, if_powers[l])
                worksheet_upc_p1db.write(row, col, raw_pout[l] + rf_loss)

                worksheet_upc_p1db_raw.write(row, 0, if_powers[l])
                worksheet_upc_p1db_raw.write(row, col, raw_pout[l])

            # Calculate the p1dB
            P1dB_index, IP1dB, OP1dB = find_p1db(if_powers, raw_pout)

            # Write the calculated P1dB information into the workbook
            # worksheet_upc_p1db_raw.write(row, 0, 'P1dB Freq (GHz)')
            # worksheet_upc_p1db.write(row, 0, 'P1dB Freq (GHz)')
            # worksheet_upc_p1db_raw.write(row + 1, col, if_powers[P1dB_index])
            # worksheet_upc_p1db.write(row, col, if_freq[P1dB_index])

            worksheet_upc_p1db_raw.write(row + 1, 0, 'IP1dB (dBm)')
            worksheet_upc_p1db.write(row + 1, 0, 'IP1dB (dBm)')
            worksheet_upc_p1db_raw.write(row + 1, col, if_powers[P1dB_index])
            worksheet_upc_p1db.write(row + 1, col, if_powers[P1dB_index])

            worksheet_upc_p1db_raw.write(row + 2, 0, 'OP1dB (dBm)')
            worksheet_upc_p1db.write(row + 2, 0, 'OP1dB (dBm)')
            worksheet_upc_p1db_raw.write(row + 2, col, raw_pout[P1dB_index])
            worksheet_upc_p1db.write(row + 2, col, raw_pout[P1dB_index] + rf_loss)

    # Return the worksheet
    return worksheet_upc_p1db, worksheet_upc_p1db_raw


##########################################################################################################
# Sweep the IF input power to see what the P1dB is
# params:   workbook, if_freq, lo_freq, rf_freq, rf_pin, lo_pin, rf_mxg, lo_mxg, specan, if_loss, lo_loss, rf_loss
# returns:  worksheet_dnc_p1db, worksheet_dnc_p1db_raw
##########################################################################################################
def rx_p1db(workbook, if_freq, lo_freq, rf_freq, rf_pin, lo_pin, rf_mxg, lo_mxg, specan, if_loss, lo_loss, rf_loss):

    # Set the mxg power levels
    rf_mxg.set_amplitude(rf_pin[0] + rf_loss)
    lo_mxg.set_amplitude(lo_pin + lo_loss)

    # Turn on the mxgs
    rf_mxg.on()
    lo_mxg.on()

    # Create the spreadsheet page
    worksheet_dnc_p1db = workbook.add_worksheet('Down_Conversion_OP1dB')
    worksheet_dnc_p1db_raw = workbook.add_worksheet('Down_Conversion_OP1dB_Raw')

    # Write the header information
    # One sheet is for the raw information and one is for the calibrated info
    row = 0
    worksheet_dnc_p1db.write(row, 0, 'RF Pin (dBm)')

    worksheet_dnc_p1db_raw.write(row, 0, 'RF Pin (dBm)')

    # Starting Column
    col = 0  # Starting at 0 instead of 1 because it immediately increments

    # Test loop
    for i in range(0, len(if_freq)):
        # Set the specan to IF frequency
        specan.set_frequency(if_freq[i] * 1e9)
        time.sleep(0.5)

        # Set the marker
        specan.set_marker(1, if_freq[i] * 1e9)
        time.sleep(0.5)

        for j in range(0, len(rf_freq)):
            # Create / Blank the lists that store the data
            if_frequencies = []
            lo_frequencies = []
            rf_frequencies = []
            rf_powers = []
            raw_pout = []

            # Reset the row for the current frequency
            row = 0

            # Increment to the next column
            col = col + 1

            # Set the rf mxg frequency
            rf_mxg.set_frequency(rf_freq[j] * 1e9)
            time.sleep(0.5)

            # Set the lo frequency
            lo_freq = (rf_freq[j] + if_freq[i]) / 4
            lo_mxg.set_frequency(lo_freq * 1e9)
            time.sleep(0.5)

            # Write the column header information
            header_info = "Pout(dBm), RF {}GHz, LO {}GHz, IF {}GHz".format(rf_freq[j], lo_freq, if_freq[i])
            worksheet_dnc_p1db.write(0, col, header_info)

            header_info_raw = "Pou_raw(dBm), RF {}GHz, LO {}GHz, IF {}GHz".format(rf_freq[j], lo_freq, if_freq[i])
            worksheet_dnc_p1db_raw.write(0, col, header_info_raw)

            for k in range(0, len(rf_pin)):
                # Sweep the RF power and set record the output power.
                # Set the RF power
                rf_mxg.set_amplitude(rf_pin[k] + if_loss)
                time.sleep(1)

                # Measure the output power
                raw_pout.append(specan.get_power(1))
                if_frequencies.append(if_freq[i])
                lo_frequencies.append(lo_freq)
                rf_frequencies.append(rf_freq[j])
                rf_powers.append(rf_pin[k])

                # Print the current Settings
                debug = "IF {}dBm, RF {}GHz, LO {}GHz, IF {}GHz".format(rf_pin[k], rf_freq[j], lo_freq, if_freq[i])
                print(debug)

            # Dump the data into the spreadsheet
            for l in range(0, len(raw_pout)):
                row = row + 1
                worksheet_dnc_p1db.write(row, 0, rf_powers[l])
                worksheet_dnc_p1db.write(row, col, raw_pout[l] + rf_loss)

                worksheet_dnc_p1db_raw.write(row, 0, rf_powers[l])
                worksheet_dnc_p1db_raw.write(row, col, raw_pout[l])

            # Calculate the p1dB
            P1dB_index, IP1dB, OP1dB = find_p1db(rf_powers, raw_pout)

            # Write the calculated P1dB information into the workbook
            # worksheet_upc_p1db_raw.write(row, 0, 'P1dB Freq (GHz)')
            # worksheet_upc_p1db.write(row, 0, 'P1dB Freq (GHz)')
            # worksheet_upc_p1db_raw.write(row + 1, col, if_powers[P1dB_index])
            # worksheet_upc_p1db.write(row, col, if_freq[P1dB_index])

            worksheet_dnc_p1db_raw.write(row + 1, 0, 'IP1dB (dBm)')
            worksheet_dnc_p1db.write(row + 1, 0, 'IP1dB (dBm)')
            worksheet_dnc_p1db_raw.write(row + 1, col, rf_powers[P1dB_index])
            worksheet_dnc_p1db.write(row + 1, col, rf_powers[P1dB_index])

            worksheet_dnc_p1db_raw.write(row + 2, 0, 'OP1dB (dBm)')
            worksheet_dnc_p1db.write(row + 2, 0, 'OP1dB (dBm)')
            worksheet_dnc_p1db_raw.write(row + 2, col, raw_pout[P1dB_index])
            worksheet_dnc_p1db.write(row + 2, col, raw_pout[P1dB_index] + rf_loss)

    # Return the worksheet
    return worksheet_dnc_p1db, worksheet_dnc_p1db_raw


##########################################################################################################
# Measure the OIP3 at all frequency combos
# params:   workbook, if_freq, lo_freq, rf_freq, if_pin, lo_pin, if_mxg_1, if_mxg_2, lo_mxg, specan, if_loss_1, if_loss_2, lo_loss, rf_loss, tone_separation_mhz
# returns:  worksheet_upconversion_ip3
##########################################################################################################
def tx_oip3(workbook, if_freq, lo_freq, rf_freq, if_pin, lo_pin, if_mxg_1, if_mxg_2, lo_mxg, specan, if_loss_1, if_loss_2, lo_loss, rf_loss, tone_separation_mhz):

    # Empty list for the measurements
    if_frequencies = []
    lo_frequencies = []
    rf_frequencies = []
    raw_pout_im_low = []
    raw_pout_tone_low = []
    raw_pout_tone_high = []
    raw_pout_im_high = []
    test_tone_separation_mhz = []

    # Set the mxg power levels
    if_mxg_1.set_amplitude(if_pin + if_loss_1)
    if_mxg_2.set_amplitude(if_pin + if_loss_2)
    lo_mxg.set_amplitude(lo_pin + lo_loss)

    # Turn on the mxgs
    if_mxg_1.on()
    if_mxg_2.on()
    lo_mxg.on()

    for i in range(0, len(if_freq)):

        for k in range(0, len(tone_separation_mhz)):
            # Calculate the IF tones
            if_low_hz = (if_freq[i] * 1e9) - (tone_separation_mhz[k] / 2 * 1e6)
            if_high_hz = (if_freq[i] * 1e9) + (tone_separation_mhz[k] / 2 * 1e6)

            # Set the IF mxg
            if_mxg_1.set_frequency(if_low_hz)
            time.sleep(0.5)

            if_mxg_2.set_frequency(if_high_hz)
            time.sleep(0.5)

            # Set the specan span
            specan.set_span(6 * tone_separation_mhz[k] * 1e6)

            for j in range(0, len(rf_freq)):
                # Set the specan frequency
                specan.set_frequency(rf_freq[j] * 1e9)
                time.sleep(0.5)

                # Set the lo frequency
                lo_freq = (rf_freq[j] + if_freq[i]) / 4
                lo_mxg.set_frequency(lo_freq * 1e9)
                time.sleep(1)

                # Calculate the tone frequencies
                rf_low_im_hz = (rf_freq[j] * 1e9) - (tone_separation_mhz[k] * 1.5 * 1e6)
                rf_low_tone_hz = (rf_freq[j] * 1e9) - (tone_separation_mhz[k] / 2 * 1e6)
                rf_high_tone_hz = (rf_freq[j] * 1e9) + (tone_separation_mhz[k] / 2 * 1e6)
                rf_high_im_hz = (rf_freq[j] * 1e9) + (tone_separation_mhz[k] * 1.5 * 1e6)

                # Set the marker to low im tone and get its power
                specan.set_marker(1, rf_low_im_hz)
                time.sleep(0.5)

                raw_pout_im_low.append(specan.get_power(1))

                # Set the marker to low main tone and get its power
                specan.set_marker(1, rf_low_tone_hz)
                time.sleep(0.5)

                raw_pout_tone_low.append(specan.get_power(1))

                # Set the marker to high main tone and get its power
                specan.set_marker(1, rf_high_tone_hz)
                time.sleep(0.5)

                raw_pout_tone_high.append(specan.get_power(1))

                # Set the marker to low im tone and get its power
                specan.set_marker(1, rf_high_im_hz)
                time.sleep(0.5)

                raw_pout_im_high.append(specan.get_power(1))

                # Append the frequency Data to the lists
                if_frequencies.append(if_freq[i])
                lo_frequencies.append(lo_freq)
                rf_frequencies.append(rf_freq[j])
                test_tone_separation_mhz.append(tone_separation_mhz[k])

    # Add a new page to the workbook
    worksheet_upconversion_ip3 = workbook.add_worksheet('Up_Conversion_IP3')

    # Write the header information
    row = 0
    worksheet_upconversion_ip3.write(row, 0, 'IF Frequency (GHz)')
    worksheet_upconversion_ip3.write(row, 1, 'LO Frequency (GHz)')
    worksheet_upconversion_ip3.write(row, 2, 'RF Frequency (GHz)')
    worksheet_upconversion_ip3.write(row, 3, 'IM Tone Separation (MHz)')
    worksheet_upconversion_ip3.write(row, 4, 'Low IM Tone Pout (dBm)')
    worksheet_upconversion_ip3.write(row, 5, 'Low Main Tone Pout (dBm)')
    worksheet_upconversion_ip3.write(row, 6, 'High Main Tone Pout (dBm)')
    worksheet_upconversion_ip3.write(row, 7, 'High IM Tone Pout (dBm)')
    worksheet_upconversion_ip3.write(row, 8, 'MAMX-011054 TX Low OIP3 (dBm)')
    worksheet_upconversion_ip3.write(row, 9, 'MAMX-011054 TX High OIP3 (dBm)')
    worksheet_upconversion_ip3.write(row, 10, 'MAMX-011054 TX Average OIP3 (dBm)')
    worksheet_upconversion_ip3.write(row, 11, 'MAMX-011054 TX Conversion Gain Low Tone (dB)')
    worksheet_upconversion_ip3.write(row, 12, 'MAMX-011054 TX Conversion Gain High Tone OIP3 (dBm)')

    # Dump the data into the worksheet
    for i in range(0, len(raw_pout_tone_low)):
        # Calculate the IP3 Values
        low_oip3 = (raw_pout_tone_low[i] + rf_loss) + (raw_pout_tone_low[i] - raw_pout_im_low[i]) / 2
        high_oip3 = (raw_pout_tone_high[i] + rf_loss) + (raw_pout_tone_high[i] - raw_pout_im_high[i]) / 2
        ave_oip3 = (high_oip3 + low_oip3) / 2

        row = row + 1
        worksheet_upconversion_ip3.write(row, 0, if_frequencies[i])
        worksheet_upconversion_ip3.write(row, 1, lo_frequencies[i])
        worksheet_upconversion_ip3.write(row, 2, rf_frequencies[i])
        worksheet_upconversion_ip3.write(row, 3, test_tone_separation_mhz[i])
        worksheet_upconversion_ip3.write(row, 4, raw_pout_im_low[i] + rf_loss)
        worksheet_upconversion_ip3.write(row, 5, raw_pout_tone_low[i] + rf_loss)
        worksheet_upconversion_ip3.write(row, 6, raw_pout_tone_high[i] + rf_loss)
        worksheet_upconversion_ip3.write(row, 7, raw_pout_im_high[i] + rf_loss)
        worksheet_upconversion_ip3.write(row, 8, low_oip3)
        worksheet_upconversion_ip3.write(row, 9, high_oip3)
        worksheet_upconversion_ip3.write(row, 10, ave_oip3)
        worksheet_upconversion_ip3.write(row, 11, raw_pout_tone_low[i] + rf_loss - if_pin)
        worksheet_upconversion_ip3.write(row, 12, raw_pout_tone_high[i] + rf_loss - if_pin)

    # Return the sheet
    return worksheet_upconversion_ip3


##########################################################################################################
# Measure the OIP3 at all frequency combos
# params:   workbook, if_freq, lo_freq, rf_freq, rf_pin, lo_pin, rf_mxg_1, rf_mxg_2, lo_mxg, specan, rf_loss_1, rf_loss_2, lo_loss, if_loss, tone_separation_mhz
# returns:  worksheet_upconversion_ip3
##########################################################################################################
def rx_oip3(workbook, if_freq, lo_freq, rf_freq, rf_pin, lo_pin, rf_mxg_1, rf_mxg_2, lo_mxg, specan, rf_loss_1, rf_loss_2, lo_loss, if_loss, tone_separation_mhz):

    # Empty list for the measurements
    if_frequencies = []
    lo_frequencies = []
    rf_frequencies = []
    raw_pout_im_low = []
    raw_pout_tone_low = []
    raw_pout_tone_high = []
    raw_pout_im_high = []
    test_tone_separation_mhz = []

    # Set the mxg power levels
    rf_mxg_1.set_amplitude(rf_pin + rf_loss_1)
    rf_mxg_2.set_amplitude(rf_pin + rf_loss_2)
    lo_mxg.set_amplitude(lo_pin + lo_loss)

    # Turn on the mxgs
    rf_mxg_1.on()
    rf_mxg_2.on()
    lo_mxg.on()

    for i in range(0, len(if_freq)):
        # Set the specan frequency
        specan.set_frequency(if_freq[i] * 1e9)
        time.sleep(0.5)

        for k in range(0, len(tone_separation_mhz)):
            # Set the specan span
            specan.set_span(6 * tone_separation_mhz[k] * 1e6)

            for j in range(0, len(rf_freq)):
                # Calculate the RF tones
                rf_low_hz = (rf_freq[j] * 1e9) - (tone_separation_mhz[k] / 2 * 1e6)
                rf_high_hz = (rf_freq[j] * 1e9) + (tone_separation_mhz[k] / 2 * 1e6)

                # Set the IF mxg
                rf_mxg_1.set_frequency(rf_low_hz)
                time.sleep(0.5)

                rf_mxg_2.set_frequency(rf_high_hz)
                time.sleep(0.5)

                # Set the lo frequency
                lo_freq = (rf_freq[j] + if_freq[i]) / 4
                lo_mxg.set_frequency(lo_freq * 1e9)
                time.sleep(0.5)

                # Calculate the tone frequencies
                if_low_im_hz = (if_freq[i] * 1e9) - (tone_separation_mhz[k] * 1.5 * 1e6)
                if_low_tone_hz = (if_freq[i] * 1e9) - (tone_separation_mhz[k] / 2 * 1e6)
                if_high_tone_hz = (if_freq[i] * 1e9) + (tone_separation_mhz[k] / 2 * 1e6)
                if_high_im_hz = (if_freq[i] * 1e9) + (tone_separation_mhz[k] * 1.5 * 1e6)

                # Need a bit of time to let the specan sweep
                time.sleep(1)

                # Set the marker to low im tone and get its power
                specan.set_marker(1, if_low_im_hz)
                time.sleep(0.5)

                raw_pout_im_low.append(specan.get_power(1))

                # Set the marker to low main tone and get its power
                specan.set_marker(1, if_low_tone_hz)
                time.sleep(0.5)

                raw_pout_tone_low.append(specan.get_power(1))

                # Set the marker to high main tone and get its power
                specan.set_marker(1, if_high_tone_hz)
                time.sleep(0.5)

                raw_pout_tone_high.append(specan.get_power(1))

                # Set the marker to low im tone and get its power
                specan.set_marker(1, if_high_im_hz)
                time.sleep(0.5)

                raw_pout_im_high.append(specan.get_power(1))

                # Append the frequency Data to the lists
                if_frequencies.append(if_freq[i])
                lo_frequencies.append(lo_freq)
                rf_frequencies.append(rf_freq[j])
                test_tone_separation_mhz.append(tone_separation_mhz[k])

    # Add a new page to the workbook
    worksheet_upconversion_ip3 = workbook.add_worksheet('Up_Conversion_IP3')

    # Write the header information
    row = 0
    worksheet_upconversion_ip3.write(row, 0, 'IF Frequency (GHz)')
    worksheet_upconversion_ip3.write(row, 1, 'LO Frequency (GHz)')
    worksheet_upconversion_ip3.write(row, 2, 'RF Frequency (GHz)')
    worksheet_upconversion_ip3.write(row, 3, 'IM Tone Separation (MHz)')
    worksheet_upconversion_ip3.write(row, 4, 'Low IM Tone Pout (dBm)')
    worksheet_upconversion_ip3.write(row, 5, 'Low Main Tone Pout (dBm)')
    worksheet_upconversion_ip3.write(row, 6, 'High Main Tone Pout (dBm)')
    worksheet_upconversion_ip3.write(row, 7, 'High IM Tone Pout (dBm)')
    worksheet_upconversion_ip3.write(row, 8, 'MAMX-011054 RX Low OIP3 (dBm)')
    worksheet_upconversion_ip3.write(row, 9, 'MAMX-011054 RX High OIP3 (dBm)')
    worksheet_upconversion_ip3.write(row, 10, 'MAMX-011054 RX Average OIP3 (dBm)')
    worksheet_upconversion_ip3.write(row, 11, 'MAMX-011054 RX Conversion Gain Low Tone (dB)')
    worksheet_upconversion_ip3.write(row, 12, 'MAMX-011054 RX Conversion Gain High Tone OIP3 (dBm)')

    # Dump the data into the worksheet
    for i in range(0, len(raw_pout_tone_low)):
        # Calculate the IP3 Values
        low_oip3 = (raw_pout_tone_low[i] + if_loss) + (raw_pout_tone_low[i] - raw_pout_im_low[i]) / 2
        high_oip3 = (raw_pout_tone_high[i] + if_loss) + (raw_pout_tone_high[i] - raw_pout_im_high[i]) / 2
        ave_oip3 = (high_oip3 + low_oip3) / 2

        row = row + 1
        worksheet_upconversion_ip3.write(row, 0, if_frequencies[i])
        worksheet_upconversion_ip3.write(row, 1, lo_frequencies[i])
        worksheet_upconversion_ip3.write(row, 2, rf_frequencies[i])
        worksheet_upconversion_ip3.write(row, 3, test_tone_separation_mhz[i])
        worksheet_upconversion_ip3.write(row, 4, raw_pout_im_low[i] + if_loss)
        worksheet_upconversion_ip3.write(row, 5, raw_pout_tone_low[i] + if_loss)
        worksheet_upconversion_ip3.write(row, 6, raw_pout_tone_high[i] + if_loss)
        worksheet_upconversion_ip3.write(row, 7, raw_pout_im_high[i] + if_loss)
        worksheet_upconversion_ip3.write(row, 8, low_oip3)
        worksheet_upconversion_ip3.write(row, 9, high_oip3)
        worksheet_upconversion_ip3.write(row, 10, ave_oip3)
        worksheet_upconversion_ip3.write(row, 11, raw_pout_tone_low[i] + if_loss - rf_pin)
        worksheet_upconversion_ip3.write(row, 12, raw_pout_tone_high[i] + if_loss - rf_pin)


##########################################################################################################
# From the provided power data determine the p1dB
# Returns the calculated IP1dB and OP1dB as well as the index
# params:   pin_list, pout_list
# returns:  p1db_index, pin_list[p1db_index], pout_list[p1db_index]
##########################################################################################################
def find_p1db(pin_list, pout_list):

    # Tracking information for the loop
    p1db_count = 0
    pout_linear = -100
    p1db_index = 0

    # Loop through the power list to find the p1dB
    for i in range(0, len(pout_list)):
        pout = pout_list[i]

        if p1db_count == 0:
            pout_linear = pout

        if p1db_count > 0:
            pout_linear = pout_linear + 1

        if pout_linear - pout > 1:
            p1db_index = i - 1
            break

        # Increment the counter
        p1db_count = p1db_count + 1

    return p1db_index, pin_list[p1db_index], pout_list[p1db_index]


##########################################################################################################
# Start of the main function
# params:   none
# returns:  none
##########################################################################################################
def main():

    # Defining the test parameters

    # IF frequency definition
    if_freq_ghz = [5.25, 5.57]

    # RF frequency definition
    rf_freq_ghz = [18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31, 32, 33, 34, 35, 36, 37, 38, 39, 40, 42, 43, 44, 45, 46]

    # Generates a list of LO frequencies calculated from the the IF and RF frequencies
    lo_freq_ghz = synth_freq_gen(if_freq_ghz, rf_freq_ghz, "lower")

    # LO power input definition
    lo_input_dbm = [13, 15, 17, 19]  # LO Input Power is typically 15 dBm, defined in the datasheet

    # RF or IF power max are both 20 dBm
    if_upc_input_dbm = 0  # Upconvert IF power input, this is to match datasheet parameters
    rf_dnc_input_dbm = 0  # Downconvert RF power input, this is to match datasheet parameters

    if_tx_p1db_start_dbm = -15  # Adjusting the starting pin to save time
    rf_rx_p1db_start_dbm = -20

    tone_separation_mhz = [20, 80, 160, 200]

    # Define the cable loss parameters
    # Need to remeasure with the new cables
    if_cable_loss_db = [0.64, 0.58]
    lo_cable_loss_db = 3
    rf_cable_loss_db = [2.01, 2.18, 2.29, 2.47, 2.34, 2.60, 2.65, 2.55, 2.73, 2.70, 2.77, 2.90, 2.90, 2.93, 3.04, 3.14, 3.30, 3.45, 3.44, 3.32, 3.58, 3.51, 3.73, 4.08, 4.36, 4.43, 4.75, 5, 5.25, 5.5]

    # Define the PCB loss parameters
    if_upc_pcb_loss_db = 0
    if_dnc_pcb_loss_db = 0
    lo_pcb_loss_db = 0
    rf_pcb_loss_db = 0

    # Cable Loss parameters for the combiner and cables
    path1_loss_5 = 0.7
    path2_loss_5 = 0.7
    path1_loss_40 = 0.7
    path2_loss_40 = 0.7

    out_cable_loss_5 = 1
    out_cable_loss_40 = 1

    # Define the maximum input values
    upc_if_max_pin = 18  # There is not a set value for this
    dnc_rf_max_pin = 18  # There is not a set value for this

    # Initialize the test equipment
    if_rf_mxg = SignalGenerator('10.13.23.221')
    lo_mxg = SignalGenerator('10.13.23.217')
    specan = SignalAnalyzer('10.13.23.222')
    if_rf_mxg2_name = 'CALIBRATION-SG.ginger.shivnet'

    # Set the default parameters on the test equipment
    if_rf_mxg.off()
    if_rf_mxg.set_amplitude(if_upc_input_dbm)
    lo_mxg.off()
    lo_mxg.set_amplitude(lo_input_dbm)

    specan.preset()
    specan.set_marker_state(1, 'ON')
    specan.set_span(.1e9)
    specan.set_rbw(1e3)

    # Set up the workbook
    current_time = datetime.datetime.now()
    spreadsheet_name_str = 'MAMX-011054_{}'.format(current_time.strftime("%Y-%m-%d_%H-%M"))
    workbook, spreadsheet_name = spreadsheet_setup(spreadsheet_name_str)

    # Create the test notes page
    test_notes = 'MAMX-011054-EVALZ Test Board Characterization'
    spreadsheet_test_info(
        workbook,
        test_notes,
        spreadsheet_name)

    # Ask what test to run
    print("Test List: \nUPCONVERT \nDOWNCONVERT \nTX_P1DB \nTX_OIP3 \nRX_P1DB \nTX_OIP3 \nRX_OIP3")
    test = input("Run what test? \n")

    if test == "UPCONVERT":

        # Run the test
        upconversion_sweep(
            workbook,
            if_freq_ghz,
            lo_freq_ghz,
            rf_freq_ghz,
            if_upc_input_dbm,
            lo_input_dbm,
            if_rf_mxg,
            lo_mxg,
            specan,
            if_cable_loss_db,
            lo_cable_loss_db,
            rf_cable_loss_db)

    if test == "DOWNCONVERT":

        # Run the test
        downconversion_sweep(
            workbook,
            if_freq_ghz,
            lo_freq_ghz,
            rf_freq_ghz,
            rf_dnc_input_dbm,
            lo_input_dbm,
            if_rf_mxg,
            lo_mxg,
            specan,
            if_cable_loss_db,
            lo_cable_loss_db,
            rf_cable_loss_db)

    if test == "TX_P1DB":
        # Adjusting the cable loss values to use the ip3 setup
        if_cable_loss_db = path1_loss_5
        rf_cable_loss_db = out_cable_loss_40

        # Calculate the range to sweep over
        step = 1  # Power Step
        if_pin = power_sweep_range(
            if_tx_p1db_start_dbm,
            if_tx_p1db_start_dbm + 30,
            step,
            upc_if_max_pin
        )

        # Run the sweep
        tx_p1db(
            workbook,
            if_freq_ghz,
            lo_freq_ghz,
            rf_freq_ghz,
            if_pin,
            lo_input_dbm,
            if_rf_mxg,
            lo_mxg,
            specan,
            if_cable_loss_db,
            lo_cable_loss_db,
            rf_cable_loss_db
        )

    if test == "RX_P1DB":
        # Adjusting the cable loss values to use the ip3 setup
        if_cable_loss_db = path1_loss_40
        rf_cable_loss_db = out_cable_loss_5

        # Calculate the range to sweep over
        step = 1  # Power Step
        rf_pin = power_sweep_range(
            rf_rx_p1db_start_dbm,
            rf_rx_p1db_start_dbm + 30,
            step,
            dnc_rf_max_pin
        )

        # Run the sweep
        rx_p1db(
            workbook,
            if_freq_ghz,
            lo_freq_ghz,
            rf_freq_ghz,
            rf_pin,
            lo_input_dbm,
            if_rf_mxg,
            lo_mxg,
            specan,
            if_cable_loss_db,
            lo_cable_loss_db,
            rf_cable_loss_db
        )

    if test == "TX_OIP3":
        # Initialize the extra test equipment
        if_rf_mxg2 = SignalGenerator(if_rf_mxg2_name)
        if_rf_mxg2.off()

        # Calculate the Loss Parameters
        if_path_1_loss_db = path1_loss_5 + if_upc_pcb_loss_db
        if_path_2_loss_db = path2_loss_5 + if_upc_pcb_loss_db
        lo_path_loss_db = lo_cable_loss_db + lo_pcb_loss_db
        rf_path_loss_db = out_cable_loss_40 + rf_pcb_loss_db

        # Run the sweep
        tx_oip3(
            workbook,
            if_freq_ghz,
            lo_freq_ghz,
            rf_freq_ghz,
            if_upc_input_dbm,
            lo_input_dbm,
            if_rf_mxg,
            if_rf_mxg2,
            lo_mxg,
            specan,
            if_path_1_loss_db,
            if_path_2_loss_db,
            lo_path_loss_db,
            rf_path_loss_db,
            tone_separation_mhz
        )

    if test == "RX_OIP3":
        # Initialize the extra test equipment
        if_rf_mxg2 = SignalGenerator(if_rf_mxg2_name)
        if_rf_mxg2.off()

        # Calculate the Loss Parameters
        rf_path_1_loss_db = path1_loss_40 + rf_pcb_loss_db
        rf_path_2_loss_db = path2_loss_40 + rf_pcb_loss_db
        lo_path_loss_db = lo_cable_loss_db + lo_pcb_loss_db
        if_path_loss_db = out_cable_loss_5 + if_dnc_pcb_loss_db

        # Run the sweep
        rx_oip3(
            workbook,
            if_freq_ghz,
            lo_freq_ghz,
            rf_freq_ghz,
            rf_dnc_input_dbm,
            lo_input_dbm,
            if_rf_mxg,
            if_rf_mxg2,
            lo_mxg,
            specan,
            rf_path_1_loss_db,
            rf_path_2_loss_db,
            lo_path_loss_db,
            if_path_loss_db,
            tone_separation_mhz
        )

    # Closes the worksheet
    workbook.close()


if __name__ == "__main__":
    main()
