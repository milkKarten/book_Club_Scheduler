# book_Club_Scheduler
Interpret survey results into groups

This program is intended for use after downloading survey results from a website as an excel file. The program removes duplicate entries. The it finds group leaders and creates groups for them based on the leaders' availability and the most popular meeting times. It then adds members to the leaders groups based on first availability. If a member cannot make any of the leaders' timeslots, the member and a leader are moved to a timeslot they can make and the members previously in that leader's group are resorted. Earlier, an average number of group members was determined from the number of total number of people and number of group leaders. Afterwards, if a group has less than the average number, group members are randomly selected from a group with more than the average members and moved into the former group. Likewise, if a group more than the average number, members are moved to other groups. The members are only able to be selected if they are able to move to the time slot of the former group. 

The only way that someone will not be added to a group that fits their own schedule is if there are no leaders available during the time(s) that they specify. The program is set to timeout after 2 seconds. If everyone is not properly sorted by this time (which happens because the sorting is psuedo-random), not everyone will be in a group. If so, simply rerun the program and everyone will be sorted properly.

This program interprets an excel file with the follow format:
1st two rows: empty
3rd row: titles of columns
1st column: First Name
2nd column: Last Name
3rd column: Write 'Yes' or 'No' depending on if one is a group leader
4th through Nth column: Write 'Available' or not for each timeslot

To use, download python, openpyxl, jdcal. Place excel spreadsheet in same folder with program and run.
