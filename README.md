# VbaChallenge

## Overview

The main objective of this analysis was to pull stock data from Excel sheets for the years 2017 and 2018 in order to find the total daily volume and return for each stock option. Daily volume is total number of shares traded throughout the day. The yearly return is the percentage difference in price starting from the beginning of each year to the end of each year. Although we could achieve this directly in Excel we can do this much faster and efficiently by enabling the use of VBA. Additionally, refactoring code to run the all of the data in one loop to improve run time and reuasability.

## Results 
After refactoring code we could see a signicant increase in run time. While my original code took around 1.2 seconds the refactored code ran in just 0.8 seconds.

![Vba_Challenge_2018](https://user-images.githubusercontent.com/110632671/186821232-c9628cfa-864f-452d-800b-c4f1f7560dd8.png)

![VbaChallenge2017 png](https://user-images.githubusercontent.com/110632671/186821279-f5557255-ce5c-4553-8fb0-8513d46275e2.png)

## Summary 
Refactoring code has many advantages such as improving the design of the code, making the code easier to understand, fixing any potential bugs, and faster coding and runtimes. However there are some disadvantages when doing so. It can be a risk when the application is very large. It can also be difficult when your set of logic may be different than the original developers making it difficult to rework and understand.
