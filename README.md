# Emily File Converter

This tool converts `.emily` files to the required output format by:

- Removing headers, footers, and metadata (everything before `[RECORDS]` and after `[END]`)
- Removing the first column (time values)
- Replacing the first column with `MOVEJ/`
- Adding commas between all columns with proper spacing
- Preserving all numbers exactly, including minus signs

## Usage

Grab the latest release in `.zip` format from the [release section](https://github.com/aptrn/emily-file-converter/releases/) and extract the archive.

**For Single Files:**

1. Drag and drop your `.emily` file onto `convert_emily.bat`
2. The converted file will be created in the same directory with `_MOVEJ.txt` suffix

**For Folders:**

1. Drag and drop a folder containing `.emily` files onto `convert_emily.bat`
2. All `.emily` files in the folder will be converted
3. Output files will be created in the same folder with `_MOVEJ.txt` suffix

## Output Format

The script converts data from this format:

```
[HEADER]
...
[RECORDS]
0.000000 +6.164479 -90.070457 +106.966484 -159.613678 +17.954603 +160.530121
0.008333 +6.164479 -90.070457 +106.966484 -159.613678 +17.954603 +160.530121
...
[END]
```

To this format:

```
MOVEJ/  +6.164479 , -90.070457 , +106.966484 , -159.613678 , +17.954603 , +160.530121
MOVEJ/  +6.164479 , -90.070457 , +106.966484 , -159.613678 , +17.954603 , +160.530121
```

## Notes

- All numbers are preserved exactly, including minus signs
- Empty lines and invalid data rows are automatically skipped
- Files without `[RECORDS]` sections are skipped with a warning
- Output files use UTF-8 encoding
- The script handles both Windows and Unix line endings

## Troubleshooting

If you get a PowerShell execution policy error, you can run this command in an administrator PowerShell window:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```
