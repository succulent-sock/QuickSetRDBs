"""
Release:  V4 - XX/XX/2025
Internal: V4
Last Updated: 07/15/2025
"""
import subprocess
import sys

def install(package):
    """Automatically runs the Python pip install command to download necessary external packages"""
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# If user does not have any packages installed, install them automatically
try:
    from tkinter import Tk
    import os
    import tkinter.filedialog
    import olefile
    from docx import Document
    import logging
    import re
    import fitz
except:
    install("python-docx")
    install("olefile")
    install("PyMuPDF")
    from docx import Document
    import olefile
    import fitz

def get_folder(message):
    """Opens a window to select a file directory"""
    Tk().withdraw()
    return tkinter.filedialog.askdirectory(title=message)

def convert_scalc_to_txt(folder):
    """Converts all Settings Basis documents into text files for data parsing"""
    basis_count = 0
    for scalc in os.listdir(folder):
        lower_name = scalc.lower()
        currentpath = os.path.join(folder, scalc)
        if ("network" not in lower_name and ("basis" in lower_name and "settings" in lower_name) or "scalc" in lower_name):
            document_name = os.path.splitext(scalc)[0]
            txt_path = os.path.join(folder, document_name + ".txt")
            if lower_name.endswith(".docx"):
                basis_count += 1
                logging.info(f"Converting {scalc}...")
                doc = Document(currentpath)
                with open(txt_path, 'w', encoding='utf-8') as f:
                    for table in doc.tables:
                        for row in table.rows:
                            row_text = '\t'.join(cell.text.strip() for cell in row.cells)
                            f.write(row_text + '\n')
            elif lower_name.endswith(".pdf"):
                # docx_path = os.path.join(folder, document_name + ".docx")
                # if not os.path.exists(docx_path):
                #     pdf_now_doc = Converter(currentpath)
                #     pdf_now_doc.convert(docx_path, start=0, end=None)
                #     pdf_now_doc.close()
                #     try:
                #         doc = Document(docx_path)
                #         all_text = [para.text.strip() for para in doc.paragraphs if para.text.strip()]
                #         with open(txt_path, 'w', encoding='utf-8') as f:
                #             f.write("\n".join(all_text))
                #         basis_count += 1
                #     except Exception as e:
                #         logging.error(f"Failed to read generated DOCX: {docx_path} — {e}")
                #         basis_count -= 1
                #     finally:
                #         if os.path.exists(docx_path):
                #             os.remove(docx_path)
                logging.info(f"Converting {scalc}...")
                currentpath = os.path.join(folder, scalc)
                document_name = os.path.splitext(scalc)[0]
                txt_path = os.path.join(folder, document_name + ".txt")
                setting_lines = []
                with fitz.open(currentpath) as pdf:
                    text = ""
                    for page in pdf:
                        text += page.get_text()
                        for line in text.splitlines():
                            if ":=" in line:
                                setting_lines.append(line)
                if setting_lines:
                    basis_count += 1
                    with open(txt_path, 'w', encoding='utf-8') as f:
                        f.write("\n".join(setting_lines))
    return basis_count

def clean_txts(folder):
    """Cleans the Settings Basis text files by parsing out the settings values in a nested dictionary."""
    # key: device name
    # value: dictionary of settings
    txt_files = {}
    for txt_file in os.listdir(folder):
        lower_name = txt_file.lower()
        if (("basis" in lower_name and "settings" in lower_name) or "scalc" in lower_name) and ".txt" in lower_name:
            logging.info(f"Cleaning {txt_file}...")
            currentpath = os.path.join(folder, txt_file)
            if "_Relay" in txt_file:
                device_name = txt_file.split("_Relay")[0]
            elif "_Settings" in txt_file:
                device_name = txt_file.split("_Settings")[0]
            elif "_SCalc" in txt_file:
                device_name = txt_file.split("_SCalc")[0]
            txt_files[device_name] = {}
            with open(currentpath, "r", encoding="utf-8") as f:
                for line in f:
                    if ":=" in line:
                        key = line.split(":=")[0].split("\t")[-1].split()[-1].strip()
                        value = line.split(":=")[1].strip().split("\t")[0].split("\r")[0]
                        if 'not used' in value.lower():
                            value = value.split(" ")[0]
                        txt_files[device_name][key] = value
    return txt_files

def delete_txts(folder):
    """Deletes the newly created text files from Settings Basis docx. They are no longer needed after program concludes."""
    paths = []
    for txt_file in os.listdir(folder):
        lower_name = txt_file.lower()
        if (("basis" in lower_name and "settings" in lower_name) or "scalc" in lower_name) and ".txt" in lower_name:
            paths.append(os.path.join(folder, txt_file))
    for file in paths:
        os.remove(file)

def clean_settings_files(folder):
    """Cleans the rdb settings text files by parsing out the settings values in a nested dictionary. """
    # key: device name
    # value: dictionary of settings by group
        # key: settings group
        # value: dictionary of settings within group
            # key: setting in rdb
            # value: current value of setting in rdb
    settings_files = {}
    for settings_file in os.listdir(folder):
        currentpath = folder + "\\" + settings_file
        if (".rdb" in settings_file.lower()):
            if "_relay" in settings_file.lower():
                device_name = settings_file.split("_Relay")[0]
            elif "settings" in settings_file.lower():
                device_name = settings_file.split("_Settings")[0]
            else:
                device_name = settings_file.split("(")[0]
            logging.info(f"Cleaning {settings_file}...")
            settings_files[device_name] = {}
            if olefile.isOleFile(currentpath):
                ole = olefile.OleFileIO(currentpath)
                # print(ole.listdir())
                for entry in ole.listdir():
                    txt_file = "/".join(entry)
                    if ".txt" in txt_file.lower() and "_" in txt_file.lower():
                        # Store settings group
                        settings_files[device_name][txt_file.split("/")[-1]] = {}
                        with ole.openstream(entry) as stream:
                            content = stream.read().decode('utf-8', errors='ignore')
                            content = content.split("\r")
                            for line in content:
                                if ',"' in line or ",'" in line:
                                    if ',"' in line:
                                        l = line.split(',"')
                                    else:
                                        l = line.split(",'")
                                    key = l[0].strip()
                                    if ',""' in line or ",''" in line:
                                        value = ''
                                    else:
                                        value = l[1].strip().strip('"').strip("'")
                                    settings_files[device_name][txt_file.split("/")[-1]][key] = value
    return settings_files

def determine_group(txt_file):
    if re.search(r"set_s\d.txt", txt_file.lower()) or re.search(r"set_\d.txt", txt_file.lower()):
        digit = re.findall(r'\d', txt_file)
        return digit[-1]
    else:
        return -1

def iterate_over_groups(all_settings_groups):
    """ For E87PG, if it's set to P, ignore the group checks for anything that is found in SET_P87.txt instead.
        If it's set to G then we do the recursive checks of each set_S1, S2 etc. """
    to_iterate_or_not_to = True
    target_setting = 'E87PG'
    for grouped_settings in all_settings_groups.values():
        current_target_value = grouped_settings.get(target_setting)
        if current_target_value:
            if current_target_value == 'P':
                to_iterate_or_not_to = False
                break
    if to_iterate_or_not_to:
        return all_settings_groups
    else:
        old_groups = all_settings_groups.keys() + []
        for group in old_groups:
            if determine_group(group) != -1:
                all_settings_groups.pop(group)
        return all_settings_groups

def compare(txt_files, settings_files):
    """Compares the dictionaries acquired from previous data cleaning. Settings Basis points are compared to Settings file points."""
    output_summary = "Settings Basis vs. Settings File Comparison Summary\n\n"
    # Iterate over all settings basis text files
    for device, basis_points in txt_files.items():
        print(f"{device:<40}{'Attribute':<20}{'Settings Basis':<40}{'Settings File':<40}")
        print(f"{'-':<40}{'-':<20}{'-':<40}{'-':<40}")
        mismatch = False
        # Find matching device settings
        matching_rdb = settings_files.get(device)
        similar_rdbs = [key for key in settings_files.keys() if device in key]
        match_found = False
        if matching_rdb:
            match_found == True
        elif len(similar_rdbs) > 0:
            matching_rdb = settings_files.get(similar_rdbs[0])
            match_found = True
        else:
            output_summary += f"Settings file for {device} was not found.\n\n"
        # Proceed if matching settings file was found
        if match_found:
            # Add device name
            output_summary += f"***** {device} *****\n\n"
            # Iterate over all points in an individual settings basis
            for key, value in basis_points.items():
                if "–" in key or "-" in key:
                    output_summary += f"{key} is listed in Settings Basis as a range of values. Please check manually.\n\n"
                    mismatch = True
                    continue

                # key: device name
                # value: dictionary of settings by group
                    # key: settings group
                    # value: dictionary of settings within group
                        # key: setting in rdb
                        # value: current value of setting in rdb
                groups_to_iterate = iterate_over_groups(matching_rdb)

                # Iterate over settings groups
                found = 0
                for settings_group in groups_to_iterate:
                    matching_point = settings_files[similar_rdbs[0]][settings_group].get(key)
                    print(f"{settings_group:<40}")
                    try:
                        print(f"{'':<40}{key:<20}{value:<40}{matching_point:<40}")
                    except:
                        print(f"{'':<40}{key:<20}{value:<40}{'':<40}")
                    if not matching_point:
                        continue
                    elif value != matching_point:
                        # Ensure numerical values are the same
                        try:
                            if float(value) != float(matching_point):
                                # Discrepancy was found
                                output_summary += f"The value of {key} in {settings_group} does not match the value of {key} in the Settings Basis.\n"
                                output_summary += f"Settings Basis: {key} := {value}\n"
                                output_summary += f"Settings File ({settings_group}): {key} := {matching_point}\n\n"
                                mismatch = True
                        except ValueError:
                            # Discrepancy was found
                            output_summary += f"The value of {key} in {settings_group} does not match the value of {key} in the Settings Basis.\n"
                            output_summary += f"Settings Basis: {key} := {value}\n"
                            output_summary += f"Settings File ({settings_group}): {key} := {matching_point}\n\n"
                            mismatch = True
                    found += 1
                if found == 0:
                    # Discrepancy was found
                    output_summary += f"{key} was not found in settings file.\n\n"
                    mismatch = True
        if not mismatch and matching_rdb:
            # Discrepancy was not found
            output_summary += "No discrepancies found!\n\n"
        print("\n")
    return output_summary

def main():
    logging.basicConfig(level=logging.INFO, format='%(message)s')

    # Select directory for Settings files + Settings Basis
    folder = get_folder("Select the directory containing your Settings Basis document(s) & Settings File(s).")
    if not folder:
        tkinter.messagebox.showerror("Error", "No valid directory selected.")
        return 1
    
    logging.info("Starting .docx .txt conversion...")
    
    # Process Settings Basis documents
    if convert_scalc_to_txt(folder) < 1:
        logging.info("No Settings Basis documents found.")
        tkinter.messagebox.showinfo("", f"No Settings Basis documents to compare.")
        return 2
    else:
        txt_files = clean_txts(folder)

        # Process rdb files
        settings_files = clean_settings_files(folder)
        if len(settings_files) < 1:
            logging.info("No rdb files found.")
            tkinter.messagebox.showinfo("", f"No rdb files to compare.")
            return 2
        
        # print(settings_files)

        notes = compare(txt_files, settings_files)

        # Remove new text files
        delete_txts(folder)

        # Output notes
        output_file = "Settings Comparison Summary.txt"
        output_file = folder + "/" + output_file
        with open(output_file, 'w') as ofile:
            ofile.write(notes)
        # Display popup window with completion message
        logging.info("Comparison Complete.")
        tkinter.messagebox.showinfo("", f"Comparison Notes saved at {output_file}")
        return 0

if __name__ == "__main__":
    main()