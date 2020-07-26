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
 sets up tiny micro server too which is possible to send requests in json file format 
 over the HTTP protocol. Those request are converted to API format of CAD software. Core code
 of application is written in python language, with use of creopyson library by Benjamin C.
 