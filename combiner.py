import os
import natsort
import argparse
import comtypes.client
from PyPDF2 import PdfFileMerger
from tqdm import tqdm



def init():
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        powerpoint.Visible = 1
        return powerpoint


def convert(target, powerpoint, formatType=32):
        if not os.path.exists(os.path.join(os.getcwd(),target,"pdf")):
                os.mkdir(os.path.join(os.getcwd(), target, "pdf"))

        files = []
        cwd = os.path.join(os.getcwd(),target)

        for filename in natsort.natsorted(os.listdir(cwd)):
                if filename.endswith('.pptx') or filename.endswith('ppt'):
                        files.append(filename)

        progress = tqdm(files, total=len(files))
        for filename in progress:
                progress.set_description("Converting : "+filename)
                newname = os.path.join(
                    cwd, "pdf", (os.path.splitext(filename)[0] + ".pdf"))
                deck = powerpoint.Presentations.Open(
                    os.path.join(cwd, filename))
                deck.SaveAs(newname, formatType)
                deck.Close()


def mergePdf(target, output, merger):
        cwd = os.path.join(os.getcwd(), target, "pdf")

        progress = tqdm(natsort.natsorted(os.listdir(cwd)),
             desc="Merged", total=len(os.listdir(cwd)))
        for filename in progress:
                progress.set_description("Merging : "+filename)
                merger.append(open(os.path.join(cwd,filename), 'rb'))
        with open(os.path.join(cwd, output+'.pdf'), 'wb') as fout:
                merger.write(fout)
        
        print("\n\n---------------------------------------------------------")
        print("SUCCESS")
        print("File saved at: "+os.path.join(cwd, output+'.pdf'))
        print("---------------------------------------------------------")





# Parse the command line inputs
parser = argparse.ArgumentParser(description='Parse the destination and output')
parser.add_argument('--folder', help='Folder name')
parser.add_argument('--output', help='Output file name')
args = parser.parse_args()

# Convert ppt to pdf
print("\nConverting ppt to pdf..")
powerpoint = init()
convert(args.folder , powerpoint)
powerpoint.Quit()

# Merge the pdf.
print("\nMerging pdfs..")
merger = PdfFileMerger(strict=False)
mergePdf(args.folder, args.output, merger)

# cwd = os.path.join(os.getcwd(),"sa","pdf")
# for files in natsort.natsorted(os.listdir(cwd)):
#         if files.endswith('.pdf'):
#                 print(files)


