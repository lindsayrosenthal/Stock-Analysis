# Stock-Analysis
VBA Stock Analysis

## Overview
Dana’s webpage and dynamic table are working as intended, but she’d like to provide a more in-depth analysis of UFO sightings by allowing users to filter for multiple criteria at the same time. In addition to the date, you’ll add table filters for the city, state, country, and shape.

## Overview

In this project and analyisis, we’ll edit, or refactor, the Stock Market Dataset with VBA solution code to loop through all the data one time in order to collect an entire dataser. Then, we’ll determine whether refactoring your code successfully made the VBA script run faster. Finally, we just want to make the code more efficient—by taking fewer steps, using less memory, or improving the logic of the code to make it easier for future users to read.



## Results

#### [Original Script found in this File](https://github.com/lindsayrosenthal/Stock-Analysis/blob/235825c6b75bab346f3ac4ab48427981d869873f/green_stocks.xlsm)
#### [Refractored Script found in this File](https://github.com/lindsayrosenthal/Stock-Analysis/blob/235825c6b75bab346f3ac4ab48427981d869873f/VBA_Challenge.xlsm)

## 2017

### Original Script:
![Original_2017](https://github.com/lindsayrosenthal/Stock-Analysis/blob/main/Resources/2017_orig.png)

### Refractored Script:
![Refract_2017](https://github.com/lindsayrosenthal/Stock-Analysis/blob/main/Resources/2017_refractored.png)


## 2018

### Original Script:
![Original_2018](https://github.com/lindsayrosenthal/Stock-Analysis/blob/main/Resources/2018_orig.png)

### Refractored Script:
![Refract_2018](https://github.com/lindsayrosenthal/Stock-Analysis/blob/main/Resources/2018_refractored.png)




## Summary
What are the advantages or disadvantages of refactoring code?

#### Disadvantages:

- A long procedure may contain the same line of code in several locations, you can change the logic to eliminate the duplicate lines.
- A logical structure may be duplicated in two or more procedures (possibly via copy & paste coding). When detected, this logic is best moved to a new function and called from the other functions.
- A complex unstructured code is usually best to split in several functions.
- Refactoring process can affect the testing outcomes.

#### Advantages:

- Logical errors easily appear in well structure code that contains nested conditionals and loops.
- In our case, using Excel flow displays program logic in a more comprehensible manner, not tied to the order that the underlying code is written.
- VBA interpretation (Excel) of code can reveal patterns that are not easy to see in the source.
