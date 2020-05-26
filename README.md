# Excel 4 Macro Generator

This program takes x86 and x64 bin files as arguments, converts them them to Excel CHAR shellcode, and adds formulas necessary to inject and execute into memory.

I could not find a good way to create the 4.0 macro sheet from python, so I have the output to a CSV file instead. You can manually run the macro code from within excel as a CSV, or if you would like a macro-enabled document, copy and paste the contents to a new XLS/XLSM document in an excel 4.0 macro sheet.

By default, outputs to 'output.csv'

# Usage

python excel4macro.py [x86payload.bin] [x64payload.bin]
