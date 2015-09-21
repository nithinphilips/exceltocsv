Excel to CSV
============
Converts Excel files to CSV or Tab delimited formats from the command-line.

This tool works by using the Excel COM API. You will need to have Excel
installed to use this.

Usage
-----
Synopsis::

    excelsheetstocsv.exe [options]  <excel-file> [<excel-file> ...]

Arguments
~~~~~~~~~

--help:
 Prints out a help message
--listformats:
 Prints a list of all supported output formats. These can be used with the
 ``--format`` argument.
--output=<outputfilename>:
 Set the output file name. This only applies if a single worksheet is selected
 using the ``--index`` or ``--pattern`` arguments.
--stdout:
 Print the output to STDOUT.
--index=<index>:
 Select the worksheet to convert by index. This is optional.
--pattern=<regex-pattern>:
 Select the worksheets to convert by using a regular expression pattern. This is
 matched against the worksheet name. Only the first matching worksheet will be
 converted. This is optional.
--format=csv|tab|<any formats supported by excel>:
 The desired output format. ``csv`` is an alias for ``xlCSVWindows`` and
 ``tab`` is an alias for ``xlTextWindows``. You may provide any other formats
 supported by Excel. Use the ``--listformats`` command to get the list of
 supported values.


If both ``--pattern`` and ``--index`` is given, ``--index`` will be used.

If neither ``--index`` nor ``--pattern`` are specified all the worksheets will
be converted. The output file names are automatically determined using the
following pattern:  Original Filename without Extension + "-" + Sheet Name + ".csv"

Example
-------
Extract the worksheet that begins with ``Data`` to ``Data.csv`` from ``Input.xlsx``::

    excelsheetstocsv --pattern=^Data --output=Data.csv Input.xlsx
