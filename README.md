# PE_class_mental_health_evaluation_program
A simple program that reads survey result.csv file and automatically analyze it and generate a report.

## Description
In a conversation with a professor from School of Physical Education at Huazhong University of Science & Technology, I found out that they are distributing surveys based on paper to their students/athletes to evaluate their mental health, and to determine their interferance based on survey evaluation results. The process of survey and evaluation is teadious and exhausting, so I came up with this simple program that would read the .csv file of survey result from online survey platform and automatically generate a report for each student.

## Installation
> Origninal project contains a QR code to the survey online. A document of the credential to opertate as the administrator of the online survey is also included. For privacy purpose, it has been remove from the project. However, the QR code to the survey is still included, as a reference to the format. This program is designed to match the format of the result from that exact survey. It most likely won't run under any other survey.
> Run the packing_to_generate_exe.bat, which will generate an final .exe file. Click on the .exe file and it will ask users to input the address where the (survey_result).csv are stored and where they want to create the (result).csv

## Feature
> When the online survey platform exports result to local PC, it will includes latest survey results as well as results from previous exports. This program will detect duplicate result from previous executing, and will not generate duplicates in its output.
> Simple UI is included, to make this program user friendly to professors at School of Physical Education, who are not familiar with PC program.
