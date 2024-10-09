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

def extract_data(unzipped_dir, out_dir):
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
    content_sets_dir = join(out_dir, "content", "sets")
    os.makedirs(content_codes_dir, exist_ok=True)
    os.makedirs(content_sources_dir, exist_ok=True)
    os.makedirs(content_quotations_dir, exist_ok=True)
    os.makedirs(content_sets_dir, exist_ok=True)

    # Construct dataframe to enable computation of simple stats like number of quotes per source, number of codes per quote, etc.
    quotes_rows = []

    # Create separate files for astro
    for code in project_json["Project"]["CodeBook"]["Codes"]["Code"]:
        code_attrs = code["attrs"]
        code_guid = code_attrs["guid"]

        with open(join(content_codes_dir, f"{code_guid}.json"), "w") as f:
            json.dump(code_attrs, f, indent=4)
    
    for source in project_json["Project"]["Sources"]["PDFSource"]:
        source_attrs = source["attrs"]
        source_guid = source_attrs["guid"]

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
                quotation["source_guid"] = source_guid

                if isinstance(quotation["Coding"], dict):
                    quotation["Coding"] = [quotation["Coding"]]

                with open(join(content_quotations_dir, f"{quotation_guid}.json"), "w") as f:
                    json.dump(quotation, f, indent=4)
                
                quotes_rows += [
                    {
                        "source_guid": source_guid,
                        "quote_guid": quotation_guid,
                        "coderef_guid": c["CodeRef"]["attrs"]["targetGUID"]
                    }
                    for c in quotation["Coding"]
                ]

    for code_set in project_json["Project"]["Sets"]["Set"]:
        set_attrs = code_set["attrs"]
        set_guid = set_attrs["guid"]

        with open(join(content_sets_dir, f"{set_guid}.json"), "w") as f:
            json.dump(code_set, f, indent=4)
    
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
    args = parser.parse_args()

    unzipped_dir = join(args.output, "unzipped")
    out_dir = args.output
    os.makedirs(out_dir, exist_ok=True)

    with zipfile.ZipFile(args.input, "r") as zip_ref:
        zip_ref.extractall(unzipped_dir)

    extract_data(unzipped_dir, out_dir)

