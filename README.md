# PDF-data-extraction
The program is able to extract information from a PDF and exhibit it as the user prefers. Please, refer to "Background" for more information.

<h1>Background</h1>

In the Vertical Farming industry, many varieties of herbs are to be harvested from different benches located in different acres. For the sake of organisation, the workers have access to a PDF file which contains all the information needed for the harvest procedure. An example of a "Harvest Label" is as follows:


The general PDF file (which contains all of the Harvest labels) is, however, hard to manipulate if the user needs to look for one (or more) specific variety. That means that the user would have to scroll up and down until the desired Harvest label is reached, which is often a time-consuming task as there is usually more than one bench with the same variety planted.

In order to solve this problem, a program was created on Python which allows the user to select only the varieties needed. And also provides additional information of how many trays[*] there are for that particular day of harvesting.

[*] Note: A tray contains the plants located in the bench. There could be up to 20 trays inside one bench.

<h1>Use of program</h1>

The user must select a date to load the Harvest labels PDF. After this, the user can select all the varieties needed and extract that specific information. It will be saved as a new PDF file.




