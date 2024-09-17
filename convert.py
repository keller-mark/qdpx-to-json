# /// script
# requires-python = ">=3.8"
# dependencies = [
#     "beautifulsoup4",
#     "lxml",
#     "bs2json",
#     "pymupdf",
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
    codes = project.find("CodeBook").find("Codes")
    code_sets = project.find("Sets")
    sources = project.find("Sources")

    out_json = join(out_dir, "output.json")
    with open(out_json, "w") as f:
        json.dump(project.to_json(), f)
    
    #print(list(sources)[0])

    # For each quotation within each source, extract the quoted region as an image file
    for source in sources:
        if source.name == "PDFSource":
            #print(source)
            pdf_guid = source["guid"]
            pdf_file = source["path"][11:]
            pdf_path = join(pdf_dir, pdf_file)

            doc = pymupdf.open(pdf_path)

            os.makedirs(join(out_dir, pdf_guid), exist_ok=True)

            selections = source.find_all("PDFSelection")
            for selection in selections:
                #print(selection)

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

                png_file = join(out_dir, pdf_guid, f"{sel_guid}.png")

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
    out_dir = join(args.output, "output")
    os.makedirs(out_dir, exist_ok=True)

    with zipfile.ZipFile(args.input, "r") as zip_ref:
        zip_ref.extractall(unzipped_dir)

    extract_data(unzipped_dir, out_dir)

