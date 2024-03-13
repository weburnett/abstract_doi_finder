# Abstract & DOI Finder

A simple tool to retrieve <abbr title="Digital Object Identifier">DOI</abbr>s and abstracts of research articles inserted as spreadsheet.

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

- * can be used to run the program on all sheets. This is also the default if no sheet range is provided.
- "1,2,3" would run the program on sheets 1, 2, and 3 in the excel file provided/found.
- "4-10" would run the program on sheets 4, 5, 6, 7, 8, 9, and 10 on the excel file provided/found.
- "10-*" would run the program from sheet 10 to the end of the excel file provided/found.

# Pre-requisites

- [Maven](https://maven.apache.org/install.html) (tested with Maven 3.9.6),
- Java (tested with Java 17.0.9),
- Place the `Publication_Abstracts_Only_Dataset_9-26-23_.xlsx` spreadsheet in `abstract_doi_finder/input/`,
