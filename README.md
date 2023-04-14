# PDF-data-extraction
The program is able to extract information from a PDF and exhibit it as the user prefers. Please, refer to "Background" for more information.

<h2>Background</h2>

In the Vertical Farming industry, many varieties of herbs are to be harvested daily from different benches located in different acres. For the sake of organisation, the workers have access to a PDF file which contains all the information needed for the harvest procedure. An example of a single "Harvest Label" is as follows:

<img src="https://user-images.githubusercontent.com/74310745/232113745-359ce038-3f73-4c75-b530-04810eda7451.png" width="600"/>

It includes all the information about the variety/varieties located inside a particular bench in the farm.

A general PDF file, which contains all of the Harvest labels, is, however, hard to manipulate if the user needs to look for one (or more) specific variety. That means that the user would have to scroll up and down until the desired Harvest label is reached, which is often a time-consuming task as there is usually more than one bench with the same variety planted.

In order to solve this problem, a program was created on Python which allows the user to select only the varieties needed. And also provides additional information of how many trays[*] there are for that particular day of harvesting.

[*] Note: A tray contains the plants located in the bench. There could be up to 20 trays inside one bench.

<h2>Use of the program</h2>

The user must select a date to load the Harvest labels PDF. After this, the user can select all the varieties needed and extract that specific information. It will be saved as a new PDF file. For example:

<img src="https://user-images.githubusercontent.com/74310745/232115320-80d1c23b-3822-48f7-99a5-b7f4c7af89e6.png" width="600"/>

Results in:

<img src="https://user-images.githubusercontent.com/74310745/232115924-58afda0d-9cda-446a-a64b-96d34329620d.png" width="400"/>

The general PDF file for this day contained 30 pages. With the program the user was able to quickly extract the two useful pages.
