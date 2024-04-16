import comtypes.client
import os
import PyPDF2

IN = 'G:\\My Drive\\Study\\UNIX\\in'
OUT = 'G:\\My Drive\\Study\\UNIX\\out'


def PPTtoPDF(inputFileName, outputFileName, formatType=32):
    powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
    powerpoint.Visible = 1

    if outputFileName[-3:] != 'pdf':
        outputFileName = outputFileName + ".pdf"
    deck = powerpoint.Presentations.Open(inputFileName)
    deck.SaveAs(outputFileName, formatType)
    deck.Close()
    powerpoint.Quit()


def combine_all_text():
    files = [f for f in os.listdir(OUT)]
    f_len = len(files)

    for index, file in enumerate(files):
        print(f'[{index}/{f_len}] Extracting from {file}')

        text = ''

        pdf_file = PyPDF2.PdfReader(os.path.join(OUT, file))  # Updated line
        num_pages = len(pdf_file.pages)

        for page_number in range(2, num_pages - 1):
            print(f'[{index}/{f_len}] [{page_number}/{num_pages}] Extracting from {file}. Page number {page_number}')
            page = pdf_file.pages[page_number]
            text += page.extract_text() + '\n\n'  # Updated line
            print(f'[{index}/{f_len}] [{page_number+1}/{num_pages}] Extracted from {file}. Page number {page_number}')

        print(f'[{index + 1}/{f_len}] Extracted from {file}')

        with open(f'G:\\My Drive\\Study\\UNIX\\txt\\{file[:-4]}.txt', 'w', encoding='utf-8') as f:
            f.write(text)


def convert():
    files = [f for f in os.listdir(IN)]
    f_len = len(files)
    for index, file in enumerate(files):
        print(f'[{index}/{f_len}] Converting {file}')
        PPTtoPDF(os.path.join(IN, file), os.path.join(OUT, f"{file[:-5]}.pdf"))  # Updated line
        print(f'[{index + 1}/{f_len}] Converted {file}')


if __name__ == '__main__':
    combine_all_text()
