# Excel 4.0 Shellcode Macro Generator

This program takes x86 and x64 shellcode bin files as arguments, converts them them to Excel CHAR shellcode, and adds formulas necessary to inject and execute into memory.

I could not find a good way to create the 4.0 macro sheet from python, so I have the output to a CSV file instead. You can manually run the macro code from within excel as a CSV, or if you would like a macro-enabled document, copy and paste the contents to a new XLS/XLSM document in an excel 4.0 macro sheet.

By default, outputs to 'output.csv'

Written for python 3.

# Usage

python excel4macro.py [x86payload.bin] [x64payload.bin]

# Creating a 4.0 Macro
In an excel sheet, the 4.0 macro option can be found by simply right-clicking on the current sheet and selecting ***Insert -> MS Excel 4.0 Macro.***

![image](https://user-images.githubusercontent.com/51035066/82890713-9b330680-9f1a-11ea-9e0f-c23b4b67bfce.png)

![image](https://user-images.githubusercontent.com/51035066/82890736-a423d800-9f1a-11ea-8937-04a98db48cc3.png)

![image](https://user-images.githubusercontent.com/51035066/82890763-ac7c1300-9f1a-11ea-91df-ba41c9fcde99.png)
