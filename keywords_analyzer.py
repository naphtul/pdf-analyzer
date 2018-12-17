import csv
import json
import re
import operator
from shutil import copyfile


with open("keywords.json", "r") as f:
    data = json.load(f)

keywords = {}
for filename, keywords_per_filename in data.items():
    keywords_as_a_list = re.split(";|,", keywords_per_filename)
    for keyword in keywords_as_a_list:
        keyword = keyword.strip(" \"")
        if keyword in keywords:
            keywords[keyword] += 1
        else:
            keywords[keyword] = 1
sorted_x = sorted(keywords.items(), key=operator.itemgetter(1), reverse=True)

keywords_list = []
for pair in sorted_x:
    if pair[0] in ["", "No Keywords were extracted from PDF", "Illegal PDF: file has not been decrypted", "Unable to extract", "xref table not zero-indexed.", "Illegal PDF: Expected object ID (6 0) does not match actual (5 0)", "Illegal PDF: Multiple definitions in dictionary at byte 0x181 for key /CreationDate"]:
        continue
    if pair[1] < 5:
        break
    kw = pair[0].replace("\"\"", "\"")
    keywords_list.append(kw)

with open("keyword.ini", "w") as f:
    f.write("[keyword]\n")
    n = 0
    for kw in sorted(keywords_list):
        n += 1
        f.write(f"keyword{n}=\"{kw}\"\n")
    f.write(f"count={n}\n")

copyfile("keyword.ini", "C:\\Users\\Public\\AppData\\Roaming\\PFU\\ScanSnap Organizer\\keyword.ini")

with open("keywords_analysis.json", "w") as f:
    json.dump(keywords, f, indent=4, ensure_ascii=False, sort_keys=True)
with open("sorted_x.json", "w") as f:
    json.dump(sorted_x, f, indent=4, ensure_ascii=False)

with open("keywords.csv", "w", newline="") as csvfile:
    writer = csv.writer(csvfile, dialect="excel")
    for keyword, keyword_count in keywords.items():
        writer.writerow([keyword, keyword_count])
