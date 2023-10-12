# Module_2_Challenge
Module 2 Challenge Due Oct 16 by 11:59pm


The first thing that I did was name the header for the cells.
Then I noticed that some of the names were bigger then the cells and I didn't like how they looked so I googled how to change a cells width. I found a solution at this website "https://learn.microsoft.com/en-  us/office/vba/api/excel.range.columnwidth"
From this I figured I could use this code to change other things about a column like number formatting.
Then I figured that the code needed was similar to the code form class so I looked at the credit_card_bonus_solutions.vbs to implement the start of my code for the for loop "i". This gave me the Stock_name and Stock_total columns.
I ran into a roadblock on how to solve ending the for loop since there could be a different number of rows for the given sheets so I googled how to grab the last row in a column. I found this website "https://www.wallstreetmojo.com/vba-last-row/" and used that information to code for column A.
After thinking about it I figured how I could take the first row <open> of the new stock and the end on the <close> by introducing another variable "j" to change and then just make the new variable the one counting and just add 1 to make it start the process for the next stock. This was possible because the code only runs the first part if the names don't equal the next cell or else it will just add the Stock_total and increase "i".
Then I just coded for finding the max and min of column K for that using what we learned in class, that is what excel does for =max/=min. Same with the max of total volume
Now I just needed to find which stocks where the values that I had already grabbed so I made an if then statement comparing those to column K and if so grabbing the stock name from column I. I wasn't too worried about having two the same since we are looking at the max and min it would be more uncommon for two to have the same but just to make sure I did a crtl+F for that max and min and there was only one of each.
From there I then added the conditional formatting of column J with the code from class posted in Slack with the color index.
I was having trouble with using the same last rows code with every new thing I did so I ended up creating 3 for the different areas I was working in which worked out. I could have probably made them all the same but it worked so I decided that it was fine.
The last thing I needed to do was to make the code run on all the worksheets in the workbook so I googled how to do that. I figured I could use another loop over everything and found this website"https://excelchamps.com/vba/loop-sheets/". With a little testing I was able to get it to work.
