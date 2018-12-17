import csv
import fnmatch
import json
import os
import re
import PyPDF2


def find_file(pattern_to_find, path_to_look_in, return_type=0, flat=False, regex=False):
    """Recursively searches in the specified path for the specified file.

    :param str pattern_to_find: The file name to look for. Regex pattern if regex=True.
    :param str path_to_look_in: The path to look in.
    :param int return_type: 0-Full, 1-Relative, 2-Filename only.
    :param bool flat: True means non-recursive. False (default) means recursive.
    :param bool regex: file_to_find is a regex expression. Default - False
    :rtype: list
    :return: An array (list) of relative path matches.
    """
    print("Trying to find {} in {}".format(pattern_to_find, path_to_look_in))
    matches = []
    for root, dirnames, filenames in os.walk(path_to_look_in):
        if regex:
            print("Using regex")
            regex_result = re.compile(pattern_to_find)
            filtered = list(filter(regex_result.match, filenames))
        else:
            filtered = fnmatch.filter(filenames, pattern_to_find)

        for filename in filtered:
            if return_type == 2:
                matches.append(filename)
            elif return_type == 1:
                rel_dir = os.path.relpath(root, path_to_look_in)
                rel_dir = "" if rel_dir == "." else rel_dir
                matches.append(os.path.join(rel_dir, filename))
            else:
                matches.append(os.path.realpath(os.path.join(root, filename)))

        if flat:
            break
    # print("Found {}".format(matches))
    return matches


def get_pdf_keywords(filepath):
    with open(filepath, "rb") as pdffile:
        try:
            pdf_toread = PyPDF2.PdfFileReader(pdffile)
            pdf_info = pdf_toread.getDocumentInfo()
        except PyPDF2.utils.PdfReadError as e:
            pdf_info = {"/Keywords": "Illegal PDF: {}".format(e)}
            print(file, "Illegal PDF: {}".format(e))
        except UnicodeEncodeError as e:
            pdf_info = {"/Keywords": "Mal-encoded PDF: {}".format(e)}
            print(file, "Mal-encoded PDF: {}".format(e))
    keywords = "Unable to extract" if pdf_info is None else pdf_info.get("/Keywords", "No Keywords were extracted from PDF")
    return str(keywords)


files_list = find_file("*.pdf", "C:\\Users\\Public\\Documents\\My ScanSnap", return_type=1)
# files_list = ["20150426115840.pdf"]
results = {}
for file in files_list:
    results[file] = get_pdf_keywords("C:\\Users\\Public\\Documents\\My ScanSnap\\{}".format(file))
keys_to_remove = []
with open("grabbed_keywords.csv", "w", newline="") as csvfile:
    writer = csv.writer(csvfile, dialect="excel")
    for file, keywords in results.items():
        try:
            writer.writerow([file, keywords])
        except UnicodeEncodeError as e:
            try:
                writer.writerow([file, "Mal-encoded keywords"])
                results[file] = "Mal-encoded keywords"
            except UnicodeEncodeError as e2:
                print("bad encoding of filename: ", file, keywords)
                keys_to_remove.append(file)

for k in keys_to_remove:
    results.pop(k)

try:
    # print(json.dumps(results, indent=4, ensure_ascii=False))
    with open("keywords.json", "w") as f:
        json.dump(results, f, indent=4, ensure_ascii=False)
except UnicodeEncodeError as e:
    print("Mal-encoded PDF: {}".format(e))
