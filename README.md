# Benchmark_Classes_File_Formatter
Formats Synergy Export file to Benchmark Classes CSV import

<H3>Updates</H3>
Version 1.0.1<br>
* Added maxRow variable - Ran into an issue when changing class sizes<br>
* Updated Kindergarten format function to change new Synergy Kindergarten export format. 

<H1>Prerequisites</h1>
Have the following Python libraries installed.
<ul><li> <a href="https://openpyxl.readthedocs.io/en/stable/">openpyxl</li>
<li><a href="https://docs.python.org/3/library/tkinter.html">tkinter</li>
<li><a href="https://pandas.pydata.org/">pandas</li></ul>


<H1>Directions</h1>

1) Download the "Benchmark Pilot Classes CSV" query from Synergy and save to downloads folder
![](https://github.com/aaronzech/images/blob/main/Screenshot_222.png)
2) Make sure the Classes.xlsx file containing the Teacher data is in the "Benchmark Files" Directory on local machine.
3) Launch the benchmark_classes_file.py script
4) Use the file picker to select the downloaded Synergy query file
5) Upload the ________ file to Benchmark's tech website.
<br></br>
![](https://github.com/aaronzech/images/blob/main/Screenshot_223.png)

<H2>Inspiration</H2>
I wanted to streamline the process from exporting the Synergy query and formatting the file manually in excel, constantly forgetting to teak a colmun header or somthing like that. So I let python do all the work to format the file into a format the Benchmark will take, and I don't need to provide any input.
