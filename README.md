# XLSB to CSV
 Python utility to convert XLSB files to CSV files

Place all the xlsb files in one directory.
From a terminal/ command prompt, run:
```
python ConvertToCSV.py path_to_xlsb_directory
```
The utility will create a new folder (by the name of csvs). All the xlsb files will be converted and stored in csvs folder.
The converted files will be named as 1.csv, 2.csv,...

The utility has a dependency on pyxlsb package.
To install the dependency, run:
```
    conda install pyxlsb
```

    or

```
    pip install pyxlsb
```
