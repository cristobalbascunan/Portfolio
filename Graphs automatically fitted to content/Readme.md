# Graphs
There is a 'Master' excel where all the information is contained, it's located in: `C:\...\master.xlsx` where all the specifications are written in separated sheets, 
one for each kind of graph except the last sheet that contains the CF, M and arrow, because they are co-dependant. Those specifications are: Name of the file output, countries, indicator, region and subregion. 


<p align="left">
<img src="https://github.com/cristobalbascunan/Portfolio/blob/e01a6931b6fa76027dc989aa33eea6a82ac5b198/Graphs%20automatically%20fitted%20to%20content/RegImg1.png" alt="Global process"\>
</p>

There are some common rules that will apply in the building of the different types of graphs but with slight differences. Let's start with the first one.




## Bar charts
1. Load specifications of graphs from master excel.
2. Depending on the 'current_publication' date, it will take the correct reportdate.
3. It will take from the database the latest report date to show the current forecasts for the latest reportdate and the historical values.
4. If the report date is equal or higher than november, it will show the next year instead of the current one. 
5. The values that may be duplicated in different tables will follow this priority: EuroArea historical data> EuroArea forecasts> Regional aggregates> Historical data >  Forecasts
6. To create the plot there are some rules in order to fit the *x* and *y* axis as much as possible to be visually attractive and informative. These rules are:
*	Get rounded limits: Convert the limits of the function rounded to the upper value (ceil or floor depending if it's positive or negative). If there are only positive or negative values the axis must start at zero.
*	Avoid overlapping: via a multiplier that will be added to the higher limit in order to create blank space to place the legend.
*	Adjust limits: uses a preset of multiples that the values should be rounded in order to get stetical visuals. This works for 6 ticks as default, but if there are some re-scaling needed (intervale of [0, 2, 4, 6, 8, 10] and the first value starts at 2.5, so it would be cut to start at 2 and have 5 ticks).
*	Adjust values to 0: if there are negative and positive values in the graphs, the interval between limits must cross the 0 value. This is made by a loop that starts by rounding the limits, and steps in order to stop if the interval between those new limits contains 0.
    *  Round to specific value depending if its minor than 1, minor than 4, minor than 10 or higher. Having different representations that will fit better what are about to being showed.
    *  Check it's higher in the negative or positive side, to start looping there.
    *  Increase a 30% the higher side and check if it would contain 0. Or if by cutting to 5 tickers it also fits.
    *  If not, it should go back to the original values and start increasing the lower side by the fixed value, with a multiplier of 10% to 100%. (i.e. it will try by adding a fixed value to the limit until reaching a % of the higher side, if not it will try over and over again until having a match, if not it will increase the opposite side too. If the end it's impossible to find a value that crosses the 0 tick, it will increase the multiplier and in the end of tries, reaching 100% that would mean a symetric plot)
