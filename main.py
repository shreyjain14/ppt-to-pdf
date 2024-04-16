import comtypes.client
import os

IN = 'G:\\My Drive\\Study\\UNIX\\abc'
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


if __name__ == '__main__':
    files = [f for f in os.listdir(IN)]
    f_len = len(files)
    for index, file in enumerate(files):
        print(f'[{index}/{f_len}] Converting {file}')
        PPTtoPDF(f'{IN}\\{file}', f'{OUT}\\{file[:-5]}')
        print(f'[{index + 1}/{f_len}] Converted {file}')
