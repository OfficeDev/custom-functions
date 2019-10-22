# License
See https://officedev.github.io/custom-functions/LICENSE

# Purpose
This folder contains two almost identical add-ins - [Upgrade App Commands](https://github.com/OfficeDev/custom-functions/blob/master/addins/upgrade/upgrade_appcmd.xml) and [Upgrade Shared Runtime](https://github.com/OfficeDev/custom-functions/blob/master/addins/upgrade/upgrade_shared.xml). Each of these two add-ins contains:
- Custom Functions
- A UI-less ribbon button handler
- A ShowTaskpane button

The only difference between these two add-ins is that Upgrade App Commands is a 'classic' add-in that uses a separate JavaScirpt runtime for each of the above three features, while Upgrade Shared Runtime uses a single, shared, runtime for all three components. 

The effect of that difference is that when the shared value in Upgrade App Commands is changed through any one of the three components that change is not visible to any of the other components, while when the shared value in Upgrade Shared Runtime is changed through any component, the change is visible to the other two components.

Additionally, these add-ins have unusual versions, which makes them suitable for testing of the upgrade experience.

# Maintainers
[zlatko-michailov](https://github.com/zlatko-michailov)
