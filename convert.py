# /// script
# requires-python = ">=3.8"
# dependencies = [
#     "beautifulsoup4",
#     "lxml",
#     "bs2json",
#     "pymupdf",
#     "pandas",
#     "numpy",
# ]
# ///
import json
import os
import zipfile
from os.path import join
from bs4 import BeautifulSoup
from argparse import ArgumentParser
from bs2json import install
import pymupdf
import pandas as pd
import numpy as np
import re

fignum_regex = r"^\d[a-zA-Z]$"

def extract_data(unzipped_dir, out_dir, smart_code_prefix = "SMART - ", exclude_text_quotes=True, exclude_fignum_codes=True, exclude_source_groups=None):
    pdf_dir = join(unzipped_dir, "sources")

    unzipped_files = os.listdir(unzipped_dir)
    
    qde_files = [ join(unzipped_dir, f) for f in unzipped_files if f.endswith(".qde") ]
    if len(qde_files) == 0:
        raise ValueError("No .qde files found")
    elif len(qde_files) > 1:
        raise ValueError("More than one .qde file found")

    qde_file = qde_files[0]

    with open(qde_file) as f:
        soup = BeautifulSoup(f, 'xml')

    project = soup.find("Project")
    sources = project.find("Sources")

    project_json = project.to_json()
    out_json = join(out_dir, "output.json")
    with open(out_json, "w") as f:
        json.dump(project_json, f, indent=4)
    
    content_codes_dir = join(out_dir, "content", "codes")
    content_sources_dir = join(out_dir, "content", "sources")
    content_quotations_dir = join(out_dir, "content", "quotations")
    content_code_groups_dir = join(out_dir, "content", "code_groups")
    content_source_groups_dir = join(out_dir, "content", "source_groups")
    os.makedirs(content_codes_dir, exist_ok=True)
    os.makedirs(content_sources_dir, exist_ok=True)
    os.makedirs(content_quotations_dir, exist_ok=True)
    os.makedirs(content_code_groups_dir, exist_ok=True)
    os.makedirs(content_source_groups_dir, exist_ok=True)

    # Create list of code for which there are corresponding smart codes
    code_names_to_ignore = []
    code_guids_to_ignore = []

    fignum_code_guids = dict() # GUID to name (e.g., "2a") mapping
    for code in project_json["Project"]["CodeBook"]["Codes"]["Code"]:
        code_attrs = code["attrs"]
        code_name = code_attrs["name"]
        if code_name.startswith(smart_code_prefix):
            code_names_to_ignore.append(code_name[len(smart_code_prefix):])
        if exclude_fignum_codes and re.match(fignum_regex, code_name) is not None:
            code_names_to_ignore.append(code_name)
            fignum_code_guids[code_attrs["guid"]] = code_name
    
    for code in project_json["Project"]["CodeBook"]["Codes"]["Code"]:
        code_attrs = code["attrs"]
        code_name = code_attrs["name"]
        if code_name in code_names_to_ignore:
            code_guids_to_ignore.append(code_attrs["guid"])

    # Construct dataframe to enable computation of simple stats like number of quotes per source, number of codes per quote, etc.
    quotes_rows = []

    # Create separate files for astro
    code_guid_to_name = dict()
    for code in project_json["Project"]["CodeBook"]["Codes"]["Code"]:
        code_attrs = code["attrs"]
        code_name = code_attrs["name"]
        if code_name.startswith(smart_code_prefix):
            # If this was a smart code, remove the prefix.
            code_name = code_name[len(smart_code_prefix):]
            code_attrs["name"] = code_name
        code_guid = code_attrs["guid"]
        code_guid_to_name[code_guid] = code_name
        
        # We need to check the code name _before_ removing any prefix.
        if code_guid not in code_guids_to_ignore:
            with open(join(content_codes_dir, f"{code_guid}.json"), "w") as f:
                json.dump(code_attrs, f, indent=4)
    
    # Sets can represent code groups (MemberCode) or source groups (MemberSource)
    source_group_name_to_member_source_guids = dict()
    for code_or_source_group in project_json["Project"]["Sets"]["Set"]:
        if "MemberCode" in code_or_source_group:
            set_attrs = code_or_source_group["attrs"]
            set_guid = set_attrs["guid"]

            with open(join(content_code_groups_dir, f"{set_guid}.json"), "w") as f:
                json.dump(code_or_source_group, f, indent=4)
        if "MemberSource" in code_or_source_group:
            set_attrs = code_or_source_group["attrs"]
            set_guid = set_attrs["guid"]

            source_group_name = set_attrs["name"]
            source_group_name_to_member_source_guids[source_group_name] = [member["attrs"]["targetGUID"] for member in code_or_source_group["MemberSource"]]

            with open(join(content_source_groups_dir, f"{set_guid}.json"), "w") as f:
                json.dump(code_or_source_group, f, indent=4)
        
    for source in project_json["Project"]["Sources"]["PDFSource"]:
        source_attrs = source["attrs"]
        source_guid = source_attrs["guid"]

        # Skip sources that are part of the excluded source groups
        skip_source = False
        if len(exclude_source_groups) > 0:
            for source_group_name in exclude_source_groups:
                if source_guid in source_group_name_to_member_source_guids[source_group_name]:
                    skip_source = True
        if skip_source:
            continue

        with open(join(content_sources_dir, f"{source_guid}.json"), "w") as f:
            json.dump(source_attrs, f, indent=4)

        # there might not be a selection in a PDF
        if "PDFSelection" not in source:
            continue

        # if there's only one selection in a single PDF,
        # `source["PDFSelection"]` is a dict and not an array
        if isinstance(source["PDFSelection"], dict):
            source["PDFSelection"] = [source["PDFSelection"]]
    
        for quotation in source["PDFSelection"]:
            if "Coding" in quotation:
                quotation_attrs = quotation["attrs"]
                quotation_guid = quotation_attrs["guid"]
                quotation_name = quotation_attrs["name"]
                quotation["source_guid"] = source_guid

                is_text_quote = ("\u00d7" not in quotation_name)
                if exclude_text_quotes and is_text_quote:
                    continue

                subfig_num = None

                if isinstance(quotation["Coding"], dict):
                    quotation["Coding"] = [quotation["Coding"]]

                # Remove codes that are to be ignored
                cleaned_codes_for_quotation = []
                for c in quotation["Coding"]:
                    code_guid = c["CodeRef"]["attrs"]["targetGUID"]
                    if code_guid not in code_guids_to_ignore:
                        cleaned_codes_for_quotation.append(c)
                    
                    if code_guid in fignum_code_guids:
                        subfig_num = fignum_code_guids[code_guid]
                    
                
                # Update the quotation with the cleaned codes, since this is what will be saved to JSON.
                quotation["Coding"] = cleaned_codes_for_quotation
                quotation["subfig_num"] = subfig_num

                with open(join(content_quotations_dir, f"{quotation_guid}.json"), "w") as f:
                    json.dump(quotation, f, indent=4)
                
                quotes_rows += [
                    {
                        "source_guid": source_guid,
                        "subfig_num": subfig_num,
                        "quote_guid": quotation_guid,
                        "coderef_guid": c["CodeRef"]["attrs"]["targetGUID"],
                        # Append code names for easier analysis.
                        "code_name": code_guid_to_name[c["CodeRef"]["attrs"]["targetGUID"]],
                    }
                    for c in quotation["Coding"]
                ]
        
    quotes_df = pd.DataFrame(data=quotes_rows)
    quotes_df.to_csv(join(out_dir, "quotes.csv"), index=True)
    
    img_dir = join(out_dir, "images")

    # For each quotation within each source, extract the quoted region as an image file
    for source in sources:
        if source.name == "PDFSource":
            pdf_guid = source["guid"]
            pdf_file = source["path"][11:]
            pdf_path = join(pdf_dir, pdf_file)

            doc = pymupdf.open(pdf_path)

            os.makedirs(join(img_dir, pdf_guid), exist_ok=True)

            selections = source.find_all("PDFSelection")
            for selection in selections:
                sel_page = selection["page"]
                page = doc.load_page(int(sel_page))

                sel_x1 = int(selection["firstX"])
                sel_x2 = int(selection["secondX"])
                sel_y1 = page.rect.y1 - int(selection["secondY"])
                sel_y2 = page.rect.y1 - int(selection["firstY"])
                sel_guid = selection["guid"]

                mat = pymupdf.Matrix(8, 8)  # zoom factor 2 in each direction

                sel_rect = pymupdf.Rect(sel_x1, sel_y1, sel_x2, sel_y2) # (x0, y0, x1, y1)
                pix = page.get_pixmap(matrix=mat, clip=sel_rect)

                png_file = join(img_dir, pdf_guid, f"{sel_guid}.png")

                with open(png_file, "wb") as f:
                    f.write(pix.tobytes("png"))
    
    print("Done")


if __name__ == "__main__":
    install()
    parser = ArgumentParser()
    parser.add_argument("--input", type=str, required=True)
    parser.add_argument("--output", type=str, required=True)
    parser.add_argument("--exclude-source-groups", nargs='*', default=[])
    args = parser.parse_args()

    unzipped_dir = join(args.output, "unzipped")
    out_dir = args.output
    os.makedirs(out_dir, exist_ok=True)

    with zipfile.ZipFile(args.input, "r") as zip_ref:
        zip_ref.extractall(unzipped_dir)

    extract_data(unzipped_dir, out_dir, exclude_source_groups=args.exclude_source_groups)

