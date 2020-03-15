This project is done to test my knowledge in UiPath in terms of using the Robotic Enterprise Framework, Orchestrator, Logging of Errors, Setting up of Config file etc.

Purpose: Generate letters easily. Simply 1) set up the template (in Word) and 2) import data (in Excel).  

The programme is created for dynamic usage (i.e. you can easily add any number of templates and "headers" and it will run properly). Therefore, you will only need to make changes to the excel and word files in order to use this programme. (No changes need to be done to the code)

p.s. you can do this with  Mail merge too, this project is done just for applying my skills)

# Getting Started
Main Files:
- 'Configuration file': 'Data/System_Config/Config.xlsx'
- 'Data for Letter Templates': 'Data/Eng_Details_Template.xlsm'
- 'Letter Templates': 'Data/Template/Normal_Engagement Letter.docx' & 'Data/Template/NSpecial_Engagement Letter.docx'

## Using the programme
1) 'Configuration file' is for providing UiPath with information on where the files are located - You may add template names and their respective directories here

2) 'Data for Letter Templates' is for inputting the data that will later be input into the Engagement letters (by matching of the "headers").

3) 'Letter Templates' are the actual letter templates that the data from 'Data for Letter Templates' will be populated in.

For more information, refer to the 'Admin' worksheets in 'Configuration file' and 'Data for Letter Templates'

### Expected Results ()
The engagement letters will be generated and saved in "New_Engagement_Letters" folder.
