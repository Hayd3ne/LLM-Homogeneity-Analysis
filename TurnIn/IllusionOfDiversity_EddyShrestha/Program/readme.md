All of the code written for our project is made to analyze the data present in 'Dataset Manager.xlsm'.



The actual dataset manager has VBA macros designed to manipulate the dataset. It's likely that Microsoft will disable macros on the dataset manager. In order to enable the macros, when looking at the file in file manager, right-click on the file, go to properties, and click "unblock" at the bottom of the general properties tab. Then click "Apply". Now, when you open the workbook, it should prompt you to enable macros at the top of the screen. All macros in the manager are designed to be used from the buttons on the "Commands" spreadsheet. Note that the "Import from File" button is no longer used (It was present from an earlier version of the project), and since changing the git repository it no longer works. 


It should be noted that running Macro-Enabled Microsoft Excel Workbooks is not supported natively on Linux, and the web version of excel does not work well with macros. In order to run, it will be necessary to set up a compatibility suite or a windows virtual machine.



As for the rest of our code, all of it is written in python. 



"embedder.py" is used by the embedding macro in the dataset manager, but could also be used to embed text on the command line. The embedding script takes two parameters, and input file path and an output file path. In our macro, this is done in a shell script as "python path/to/embedder.py path/to/input.json path/to/output.json". Obviously, this requires python to be on PATH and a computer that can run python.



Both of our analysis files are saved as Jupyter Source Files (".ipynb"). There are multiple ways to run these files, even in VSCode, but we used Google Collaboratory as an IDE for this project. You can simply open the files in Colab to run them.
.

