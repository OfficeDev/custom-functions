# License
See https://officedev.github.io/custom-functions/LICENSE

# Purpose
This folder contains four add-ins - [OfficeJs1-Sample-Excel](https://github.com/OfficeDev/custom-functions/blob/master/addins/officejs2/agave/manifest-v1.xml), [OfficeJs2-Sample-Excel](https://github.com/OfficeDev/custom-functions/blob/master/addins/officejs2/agave/manifest-v2.xml),[OfficeJs1-Sample-Word](https://github.com/OfficeDev/custom-functions/blob/master/addins/officejs2/agave/Word-manifest-v1.xml), [OfficeJs2-Sample-Word](https://github.com/OfficeDev/custom-functions/blob/master/addins/officejs2/agave/Word-manifest-v2.xml). All of the four add-ins contains a "show taskpane button" which could call basic RichApi to make sure office.js is initilized successfully.

The difference between V1 and V2 is: 
In V1 addin, https://appsforoffice.microsoft.com/lib/beta/hosted/office.js is referenced; 
In V2 addin, https://appsforoffice.microsoft.com/lib/beta/hosted/word.js or https://appsforoffice.microsoft.com/lib/beta/hosted/excel.js is referenced for perfomance optimization. 

# Maintainers
[JiayuanL](https://github.com/JiayuanL)
[shaofengzhu](https://github.com/shaofengzhu)