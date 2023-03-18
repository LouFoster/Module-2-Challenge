## Overview of Project
(As captured from Data Bootcamp Module 2) 

As a Data Analyst, I am helping Steve (the Client) analyze an entire dataset. In addition, to do a little more research for Steve’s parents, the client wants to expand the dataset to include the entire stock market over the last few years.

Steve has given me, the Data Analyst, an Excel file containing the stock data he wants to be analyzed. As Data Analyst, I will be using an extension to Excel, built to automate tasks. Previously, Excel helped us start thinking about data differently; now, VBA will help us start thinking about how to analyze that data programmatically. 

By the end of this module, as assigned Data Analyst, I will complete the following :

•	Create a VBA macro that can trigger pop-ups and inputs, read, and change cell values, and format cells.

•	Use for loops and conditionals to direct logic flow.

•	Use nested for loops.

## Purpose
(This module is built around a project that mirrors a real-world scenario that would require data analysis and visualization.) 
This new assignment consists of one technical deliverable and a written report to deliver your data analyst results. As such, I will submit the following:

•	Deliverable 1: Refactor VBA code and measure performance.
    o	This deliverable will include an updated workbook and a folder with PNGs of the pop-ups with script run time.

•	Deliverable 2: A written analysis of your results (README.md).

In addition, this project’s purpose is to compare the original VBA script to the refactored script and show improvement in runtime speed.

## Results

##Deliverable 1: Analysis of " Refactor VBA Code and Measure Performance

As the assigned Data Analyst, I will use my knowledge of VBA and the starter code provided in this Challenge to refactor the Module2_VBA_Script to loop through the data once and collect all the information. In addition, the goal is that the refactored code should run faster than it did in this module.

#Sub AllStocksAnalysisRefactored() for 2017
![image](https://user-images.githubusercontent.com/117233641/226089814-8a2bf9f9-9fce-404d-828a-798c04b79920.png)
 
Sub AllStocksAnalysisRefactored() for 2018
![image](https://user-images.githubusercontent.com/117233641/226089832-04c37c63-34d4-48aa-9290-fd51f4f80c35.png)
 
 
VBA Code for Sub AllStocksAnalysisRefactored()as per Challenge instructions
 ![image](https://user-images.githubusercontent.com/117233641/226089856-5fba5322-d0c8-477b-a198-beeddeecdb75.png)

![image](https://user-images.githubusercontent.com/117233641/226089867-4e046846-7845-4a91-9b17-945f7ba387d4.png)

![image](https://user-images.githubusercontent.com/117233641/226089877-aec70e28-e332-4725-a412-7dfb0e97bf1a.png)
 
 ![image](https://user-images.githubusercontent.com/117233641/226089891-f05a9989-e759-43b7-9be4-8c5ace0bdea7.png)

 



##Deliverable 2: Written Analysis of Results (20 points)

Overview of Project: Explain the purpose of this analysis.

1.	Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.

Running our 2017 and 2018 data stock analysis results in seeing an elapsed run time for each year (see Screen Shots below for data visualization). 

•	In addition, regarding stock performance, there were more positive returns in 2017 compared to 2018. 
•	Further analysis shows that five stocks had a decreased total daily volume in 2018, whereas the remaining 7 had an increased total daily volume in 2018.

#Sub AllStocksAnalysisRefactored() for 2017
![image](https://user-images.githubusercontent.com/117233641/226089904-8b3e5e6e-d7a1-4002-b468-27e01021a364.png)
 
Sub AllStocksAnalysisRefactored() for 2018
 ![image](https://user-images.githubusercontent.com/117233641/226089910-a04c6d18-a5bb-4c0c-9583-4d3ecad285c5.png)



##Summary: In a summary statement, address the following questions.

1.	What are the advantages or disadvantages of refactoring code?
Improving or updating the code without changing the software’s functionality or external behavior of the application is known as code refactoring. 

Advantages of refactoring code

•	faster runtime, requiring fewer steps, less memory, and easier code readability for future users since it only loops through all the data one time.

•	more adaptability as it can handle larger datasets resulting in a measurable greater efficiency.

•	Due to the step-by-step structure of the code, logical errors are more easily recognized, especially code containing loops. 

Disadvantages of refactoring code

•	Complex code is usually best to split into several functions in order to spot errors and or coding issues.

•	refactoring code requires a meaningful understanding of the original in order to achieve maximum optimization.


2.	How do these pros and cons apply to refactoring the original VBA script

•	As per our results, the refactored code allowed for a faster runtime; see attached screenshots.

•	A clean organized code with the enhanced ability to make changes, which are more easily manageable.

•	Refactoring code adds additional time to produce a final product as it most would require additional time spent working to optimize the code. 


