# Automation of CAD Model Creation 
This application is dedicated to creation of CAD master model with use of 
creoson application and creopyson API. Implementation of this application
happened in Kraussmaffei Technologies company, which focuses on 
production of injection moulding machines. Application automates manual 
process of creation CAD master model. This process is based on assembling 
of models in to master model according to information from ERP system. 

## Table of Contents
* [Workflow of application](#workflow-of-application)
* [Technologies](#technologies)
* [Dependencies and Folders](#dependencies-and-folders)

## Workflow of application 
 
* User makes sure creoson application is initialized and set up.
  Then is application allowed to run.
* User selects source master model and order number for automation process.
* Automation loads bill of material from ERP system. After it takes control
  control over CAD software CREO Parametric with use of creoson API and creopyson library.
* CAD master model is prepared and correct model groups are assembled to it from database system.  
* Feedback of process is delivered to user, to save postprocess time of finalization of CAD master model.

## Technologies

 This project uses Creoson application from Simplified Logic inc. This application
 sets up tiny micro server through which is possible to send requests in json file format 
 over the HTTP protocol. Those request are converted to API of CAD software. Core code `KM_Assembly_Automation.py`
 is written in python language, with use of creopyson library by Benjamin C. This core code uses procedural approach.

## Dependencies and Folders
Except direct dependency on performance of Creoson application, automation of mastermodel has
several more dependecies.
* Core application `KM_Assembly_Automation.exe` is dependent to folder it is located in. 
* In `DatabaseFolder` 
is located `mastermodels_database.xlsx` file, which consists list of source models of injection moulding 
machines and their properties (like clamping unit size, injection unit size, number of plastification units, 
special signs etc.) 
* `DeleteExclude` folder contains sub-folders with name of machine types manufactured by company. `DeleteExclude.csv` 
file is in every sub-folder, and consists information of assembly groups which will be spare in process of automation 
due to low data quality of its CAD models (such as incorrect naming conventions).
* `FeedbackFolder` is folder where information of process are stored. After every automation there are screenshots
of model groups which should be in CAD model according to bill of material list, but various kind of errors those
CAD models had not been assembled to CAD master model.
* `IconPictures` folder contains icons and picture which are used in graphical user interface of 
applications
* `ErpBom` folder is used to store bill of material from ERP system



 
## Sources
* creopyson library - https://github.com/Zepmanbc/creopyson
* creoson application - https://github.com/SimplifiedLogic/creoson
