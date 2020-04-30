# Excel 4 Macro Generator

This program takes x86 and x64 bin files as arguments, converts them them to Excel CHAR shellcode, and adds formulas necessary to inject and execute into memory.

By default, outputs to 'output.csv'

# Example

python excel4macro.py x86payload.bin x64payload.bin
