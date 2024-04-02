# Abstract & DOI Finder

A simple tool to retrieve <abbr title="Digital Object Identifier">DOI</abbr>s and abstracts of research articles inserted as spreadsheet.

## How-to Use

- First, we need some set-up:
    - [Prepare a spreadsheet](#sheet-requirements),
    - If you don't have an account on [github](https://github.com/), [create one](https://github.com/signup),

- We will now create a copy of the program:
    - Go to <https://github.com/popbr/abstract_doi_finder/fork>, to "fork" (that is, create a copy of) our repository (that is, the program code),
    - Leave everything by default, and click on "Create fork":
        ![](how_to/fork_in_github.png)
    - You should now see your own fork: you can make sure by checking that the address bar is of the form `https://github.com/<your-username>/abstract_doi_finder`, where `<your-username>` is your username (`aubertc` in the example below). 

- Now we will delete the spreadsheet loaded by default.
    Click on the "abstract_doi_finder" folder:
    ![](how_to/naviguate_in_github_1.png)
    then click on the "input" folder, and on the "test_input.xlsx" file. On the right of "Go to file", click on the three dots, then on "Delete file":
    ![](how_to/naviguate_in_github_4.png)
    Finally, click on "Commit changes…" twice:
    ![](how_to/naviguate_in_github_5.png)

- We will now upload our own spreadsheet.
    Make sure you are still in the "input" folder, and click on the "Upload files" button hidden under the "Add file" button:
    ![](how_to/naviguate_in_github_2.png)
    Upload your spreadsheet, and click on "Commit changes"
    
- Now, we will execute the program on our spreadsheet.
    Click on "action"



## Sheet Requirements




_A guide on how to use this tool will follow._

```
cd abstract_doi_finder/
mvn compile
mvn exec:java -Dexec.mainClass="popbr.AbstractDoiFinder" -Dexec.args="Publication_Abstracts_Only_Dataset_9-26-23.xlsx 1,2,3"
```

where `Publication_Abstracts_Only_Dataset_9-26-23.xlsx` is the name of the spreadsheet placed in the `abstract_doi_finder/input/` folder.

and

where '1,2,3' is the range/sheets you want to run the program on. Please separate the values with commas, exclude spaces, or follow the examples below.

The range can be provided in a multitude of ways including:

- "*" can be used to run the program on all sheets. This is also the default if no sheet range is provided.
- "1,2,3" would run the program on sheets 1, 2, and 3 in the excel file provided/found.
- "4-10" would run the program on sheets 4, 5, 6, 7, 8, 9, and 10 on the excel file provided/found.
- "10-*" would run the program from sheet 10 to the end of the excel file provided/found.

# Pre-requisites

- [Maven](https://maven.apache.org/install.html) (tested with Maven 3.9.6),
- Java (tested with Java 17.0.9),
- Place the `Publication_Abstracts_Only_Dataset_9-26-23_.xlsx` spreadsheet in `abstract_doi_finder/input/`,
