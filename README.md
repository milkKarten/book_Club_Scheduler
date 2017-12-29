# book_Club_Scheduler
Interpret survey results into groups

This program is intended for use after downloading survey results from a website as an excel file. The program removes duplicate entries. The it finds group leaders and creates groups for them based on the leaders' availability and the most popular meeting times. It then adds members to the leaders groups based on first availability. Earlier, an average number of group members was determined from the number of total number of people and number of group leaders. Afterwards, if a group has less than the average number, group members are randomly selected from a group with more than the average members and moved into the former group. The members are only able to be selected if they are able to move to the time slot of the former group. 

Since the groups are chosen on popularity, this program works for most cases. The following issue is typically avoided by using the most popular times. If everyone chooses low number of availabilities, this program will not always work. Future work will be done to be able to switch the group leaders timeslots if there are people who cannot make any of the current timeslots. 

This program interprets an excel file with the follow format:
1st two rows: empty
3rd row: titles of columns
1st column: First Name
2nd column: Last Name
3rd column: Write 'Yes' or 'No' depending on if one is a group leader
4th through Nth column: Write 'Available' or not for each timeslot
