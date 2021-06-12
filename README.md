# excel-automation-template

# About

This is a collection of Python automation scripts to help me with the data analysis for my research project on [Cognitive Load](https://github.com/vondreii/cognitive-load).

Each participant was required to play two versions of a game. For each version, they needed to play the game three times. In total, one participant had 6 excel files. In total there were 40 participants. 40*6 = 240 excel files. 

![img](/Images/cmd.PNG)

### Data Analysis Script 
This script does simple analysis and maths calculations on the data. In order to speed up the process, the data analysis script was based off Sora Khan's Data Analysis Script: [https://github.com/sorakhan/data-analysis-script](https://github.com/sorakhan/data-analysis-script)

The script has since been modified to include new summary tables and maths calculations. It calculates things such as:
- Finding the average reaction times at specific times in the game
- Calculating the total crashes within a specific timeframe during the gameplay (eg, between 10-20 seconds)
- Calculating the Standard Deviation and Averages of the player's response times, and DRT misses

![img](/Images/excelDataAnalysis.png)

### Template 
I also created a template which can be used to create new Python scripts to read excel files.

### Clean Data
This Script deletes certain columns in the raw excel file that are not needed for the analysis.

