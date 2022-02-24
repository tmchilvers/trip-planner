# trip-planner
**|**  v.01.01 **|  Author:** [Tristan Chilvers](https://www.tristanchilvers.com/) **| License:** [MIT License](LICENSE) **|**

<br>
<p align="center"><b>A SET OF GOOGLE CUSTOM FUNCTIONS FOR A DYNAMIC TRIP PLANNER SPREADSHEET</b></p>

                      [DOWNLOAD](#create-your-own-vacation-planner)

---

<br>

For my vacations, I often times create a trip outline with Google Sheets. These are rough calendar outlines with every event *manually* inputted and adjusted to fit into the calendar - way too slow and tedious!
<br><br>

While there are many great vacation planners online, I wanted to focus on two qualities for my own trip planner: 
<br><p align="center">
  <ins><b>Automated</ins></b> and <ins><b>Dynamic</ins></b> Calendar & Finance Functionality
</p>

<br>

---

# SHEETS
<p align="center"><b>DETAILED EXPLAINATION FOR EACH SHEET</b></p>

    [• TRIP DETAILS](#trip-details)
    [• TRAVELERS](#travelers)
    [• EVENTS](#events)
    [• CALENDAR](#calendar)
    [• FINANCE](#finance)

<br>

## TRIP DETAILS

The *Trip Details* sheet handles global details relevant to the trip:
- Trip Dates, Need Passport, Budget, etc.

You only need to input data into the **white cells** (the grey cells are calculated and displayed for your convenience)

Whenever you update the **start date** and **end date** cells, it will automatically call a script to update the calendar sheet. This is one example of a core feature that performs an *automated* and *dynamic* task - the calendar will automatically resize itself and update the dates to outline the duration of your trip.

<p align="center"><b><ins>NOTE: DO NOT EDIT EVENTS UNTIL AFTER THE CALENDAR HAS FINISHED UPDATING</ins></b></p>

<p align="center">
<img src="/../main/gifs/trip-details-date-change.gif" alt="trip-details-date-change-preview" style="width:45%;"/>
  <br><br>
<img src="/../main/gifs/calendar-date-change.gif" alt="calendar-date-change-preview" style="width:75%;"/>
</p>

<br>

If you select **Yes** for **Need Passport?**, it will mark **red** next to each traveler in the traveler sheet who does not have a passport.

<br>

<p align="center">
<img src="/../main/images/passport-check.png" alt="passport-check-preview" style="width:100%;"/>
</p>

<br>

Depending on your selection for **Miles** or **km** for, it will update the column name under the *Events Sheet* to correspond with your choice of units.

<br>

<p align="center">
<img src="/../main/images/miles_or_km.png" alt="calendarPreview" style="width:19%;"/>
<img src="/../main/images/miles_or_km_02.png" alt="calendarPreview" style="width:30%;"/>
<img src="/../main/images/miles_or_km_03.png" alt="calendarPreview" style="width:30%;"/>
</p>

<br>

## TRAVELERS

The *Travelers* sheet handles each traveler's information:
- Name, Email, Phone, etc.


<br>

<p align="center">
<img src="/../main/images/passport-check.png" alt="travelers-preview" style="width:100%;"/>
</p>

<br>

The amount of money each traveler has paid in relationship to the trip's budget is displayed under the *Finance Sheet*.

<br>

<p align="center">
<img src="/../main/images/traveler-finance-preview.png" alt="finance-preview" style="width:60%;"/>
</p>

<br>

## EVENTS

The *Events Sheet* is the backbone of this software - the *Calendar Sheet* and *Finance Sheet* will adjust according to the data inputted here.

This planner is structured by **events**. An **event** is a row containing all data relevent to that event on your trip:

- Event name, date, time, cost, split charge, etc.

<br>

<p align="center">
<img src="/../main/images/Events_Preview.png" alt="events-preview" style="width:100%;"/>
</p>

<br>

Each event is labeled as a type, which is then categorized and displayed on the *Calendar Sheet* and *Finance Sheet*:

- **Transportation, Lodging, Food, Activity, Shopping, and MISC**

<br>

Once you have inputted your events, you can update the calendar by clicking the **CLICK TO UPDATE CALENDAR** button on the top-left corner.

<p align="center"><b><ins>NOTE: DO NOT EDIT EVENTS UNTIL AFTER THE CALENDAR HAS FINISHED UPDATING</ins></b></p>

<p align="center">
<img src="/../main/gifs/update-calendar-events.gif" alt="update-calendar-events-preview" style="width:45%;"/>
  <br><br>
<img src="/../main/gifs/update-calendar-result.gif" alt="update-calendar-result-preview" style="width:75%;"/>
</p>

## CALENDAR

The *Calendar Sheet* outlines each day of your trip into a clean, dynamic layout of your events.

Each event is of a certain type (Transportation, Lodging, etc.) and is displayed here by the type's corresponding color (e.g. Red for Transportation).

It showcases the distance travelled each day (useful for Road Trips), along with the calculated cost of gas.

<br>

<p align="center">
<img src="/../main/images/Calendar_Preview.png" alt="calendar-preview" style="width:75%;"/>
</p>

<br>

You can also update the calendar by clicking the **CLICK TO UPDATE CALENDAR** button on the top-left corner.

<p align="center"><b><ins>NOTE: DO NOT EDIT EVENTS UNTIL AFTER THE CALENDAR HAS FINISHED UPDATING</ins></b></p>

<p align="center">
<img src="/../main/gifs/update-calendar.gif" alt="update-calendar-preview" style="width:35%;"/>
  <br><br>
<img src="/../main/gifs/update-calendar-result.gif" alt="update-calendar-result-preview" style="width:75%;"/>
</p>

## FINANCE

The *Finance Sheet* outputs the financial information relevant to the trip:
- Budget vs Final Cost, where was the money spent (transportation, food, etc.), cost of each traveler, etc.

<br>

<p align="center">
<img src="/../main/images/finance-preview.png" alt="finance-preview" style="width:75%;"/>
</p>

<br>

---

# CREATE YOUR OWN VACATION PLANNER
Make a copy of this planner template **(check for updates!)**:
- [Trip Planner (TEMPLATE)](https://docs.google.com/spreadsheets/d/1B_dnR0HjFd-yClSEWd2_BwNsqBRuEJXZhJbATznj11U/edit?usp=sharing)

Or you can copy/edit this code directly in your project by going to **Tools > Script editor** and
pasting [this code](trip.js) into `trip.gs`.

