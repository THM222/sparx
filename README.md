# sparx

Application that takes Sparx data as input and outputs a PowerPoint presentation summarising some data.

## User Instructions

1. Download a [Student Activity Report](https://www.sparxmaths.uk/teacher/reports) from Sparx
2. Open the Application
3. Click `Choose File` and select your newly downloaded Sparx student activity data
4. Click `Choose Directory` and select the folder where you would like the output PowerPoint to be saved
5. Click `Run`
6. A popup will appear once the data has been processed. Click OK. A File Explorer will open to where the PowerPoint was saved.
7. There will be a folder named `sparx_<YYYYmmdd>`, where `YYYYmmdd` is today's date e.g. 13 Jan 2025 would be `20250113`.
8. Inside the `sparx_<YYYYmmdd>` folder will be a file named `sparx_leaderboard.pptx`!

### Advanced options

Click Advanced on the UI to further customise the output.

| Option | Details |
|--------|------------------|
| Weeks to process| Number of weeks of data to process. Defaults to 2. |
| XP Boost Top N | Number of students to show on the XP Boost leaderboard. Defaults to 10. |
| Independent Learning Top N | Number of students to show on the Independent Learning leaderboard. Defaults to 10. |
| Independent Learning minimum minutes | The minimum number of minutes of independent learning to make the leaderboard. Defaults to 20 mins |
| Process year group data | Create slides for the year group completion data. |
| Process maths class data | Create slides for the completion by maths class data. |
| Process registration group data | Create slides for the completion by form class data. |
| Process independent learning | Create slides for the independent learning student leaderboard. |
| Process XP Boost | Create slides for the XP Boost student leaderboard. |



## Developer Instructions

Prerequisites:
- git installed
- python3-tk installed [stackoverflow](https://stackoverflow.com/a/74607246)

To setup and run locally:
1. `git clone https://github.com/THM222/sparx.git`
2. `cd sparx/ui`
3. `python -m venv venv`
4. `source venv/bin/activate`
5. `pip install -r requirements.txt`
6. `python -m main`
