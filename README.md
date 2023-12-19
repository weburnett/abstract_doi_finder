# Abstract & DOI Finder

A simple tool to retrieve <abbr title="Digital Object Identifier">DOI</abbr>s and abstracts of research articles inserted as spreadsheet.

_A guide on how to use this tool will follow._

```
cd abstract_doi_finder/
mvn compile
mvn exec:java -Dexec.mainClass="popbr.AbstractDoiFinder" -Dexec.args="Publication_Abstracts_Only_Dataset_9-26-23.xlsx"
```

where `Publication_Abstracts_Only_Dataset_9-26-23.xlsx` is the name of the spreadsheet placed in the `abstract_doi_finder/input/` folder.

# Pre-requisites

- [Maven](https://maven.apache.org/install.html) (tested with Maven 3.9.6),
- Java (tested with Java 17.0.9),
- Place the `Publication_Abstracts_Only_Dataset_9-26-23_.xlsx` spreadsheet in `abstract_doi_finder/input/`,
