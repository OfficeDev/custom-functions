# License
See https://officedev.github.io/custom-functions/LICENSE

# Purpose
This add-in is used to test Shared Runtime with Custom Functions, Taskpane and App Commands.
This add-in contains:
- Custom Functions
- A ShowTaskpane button
- A UI-less ribbon button handler

This Addin contains three components:
1. Taskpane: Contains six buttons, [Set/Read]Data - to set or get the value from the shared global variable g_sharedAppData',  Show/Hide - to show or hide the taskpane and [Get/Set]StartupState - to get current Runtime startup behavior or set current Runtime Startup state to Load or None.
2. UI-less button handler: To update the value of shared global variable g_sharedAppData' to '2019'.
3. Custom Functions: Contain functions like [Set/Get]Value - to set or get the value from the shared global variable g_sharedAppData', Show/Hide - to show or hide the taskpane, [Get/Set]StartupBehavior - to get or set current Runtime startup behavior, [Get/Set]RangeValue set/get value for a Range.

# Office.js
Please reference the release version of office.js
```html
		<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

# Manifest Changes for Shared App
Shared App is an HTML page that will be used for taskpane, app command and custom functions. To use Shared App feature, please update the manifest. For example, this addin's manifest added the following runtime definition under the Host element.
```xml
        <Runtimes>
          <Runtime resid="OEP.SharedRuntime.Url" lifetime="long" />
        </Runtimes>
```
It declared one runtime and specify the runtime's url using resid.

Then it declared that the custom function to use the shared runtime by
```xml
        <Page>
          <SourceLocation resid="OEP.SharedRuntime.Url"/>
        </Page>
```
As it uses the same resid as the one declared in the runtime, the custom function will use the shared runtime.

It also declared that the app command also use the shared runtime by
```xml
        <FunctionFile resid="OEP.SharedRuntime.Url" />
```
As it uses the same resid as the the one declared in the runtime, the app command will use the shared runtime.

For taskpane, 
```xml
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="OEP.SharedRuntime.Url" />
        </Action>
```
As the ShowTaskpane actioin uses the same resid as the one declared in the runtime, the taskpane will use the shared runtime.

# Manifest Changes for Enable/Disable Ribbon API
The <Enabled> element under <Control> is used to control the default/initial enable or disable state of the ribbon control, by default, the control is enabled.
```xml
      <Control>
        ......
        <Action>
          ........
        </Action>
        <Enabled>false</Enabled>
      </Control>
```

# Preview API
## Contextual Tabs related APIs

### Contextual Tab Json defination
```json
      {
        "actions":[
          {
            "id":"executeWriteData",
            "type":"ExecuteFunction",
            "functionName":"writeData"
          },
          {
            "id":"showTaskpaneResCFSample",
            "type":"ShowTaskpane",
          }
          ],
          "tabs":[
          {
            "id":"CtxTab1",
            "label":"CtxTab1_label",
            "visible":true,
            "groups":[
              {
                "id":"CustomGroup111",
                "label":"Group11Title",
                "icon":[
                {
                  "size":16,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                },
                {
                  "size":32,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                },
                {
                  "size":80,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                }
                ],
                "controls":[
                {
                  "type":"Button",
                  "id":"CtxBt111",
                  "enabled":false,
                  "icon":[
                {
                  "size":16,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                },
                {
                  "size":32,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                },
                {
                  "size":80,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                }
                          ],
                          "label":"STP_CtxBt111",
                          "toolTip":"Btn111ToolTip",
                          "superTip":{
                "title":"Btn111SupertTipeTitle",
                "description":"Btn111SuperTipDesc"
                          },
                          "actionId":"showTaskpaneResCFSample"
                        },
                        {
                          "type":"Button",
                          "id":"CtxBt112",
                          "enabled":true,
                          "icon":[
                          {
                  "size":16,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                },
                {
                  "size":32,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                },
                {
                  "size":80,
                  "sourceLocation":"https://officedev.github.io/custom-functions/addins/CFSample/Images/Button32x32.png"
                }
                          ],
                          "label":"ExeFunc_CtxBt112",
                          "toolTip":"Btn112ToolTip",
                          "superTip":{
                "title":"Btn112SupertTipeTitle",
                "description":"Btn112SuperTipDesc"
                          },
                          "actionId":"executeWriteData"
                        }
                      ]
                    }
                  ]
                }
              ]
            };
```

### Create Contextual Tabs
```js
      // Contextual Tab defination in JSON blob
      var ribbonTabDefinition = document.getElementById('ctxblob').value;
      // Request to create the Contextual Tab
      Office.ribbon.requestCreateControls(JSON.parse(ribbonTabDefinition))
      .then(function(result) {
        log('Success:' + JSON.stringify(result));
      })
      .catch(function (error) {
        log('Error:' + JSON.stringify(error));
      });
```
### Update Contextual Tabs visibility
```js
      // Target Contextual Tab Id defined in Json blob
      var tabIds = document.getElementById('ctxid').value.split(';');
      var btn = {};
      var commandtabs = [];

      for (var i = 0; i < tabIds.length; i++) {
        var tabId = tabIds[i];
        if (tabId != "") {
          // Update the Visible state accordingly
          commandtabs.push({id: tabId, visible: Boolean(visible), controls: [btn]});
        }
      }

      var data = {tabs: commandtabs};
      // Send the request to update the Contextual Tab
      Office.ribbon.requestUpdate(data)
      .then(function(result) {
        log('Success:' + JSON.stringify(result));
      })
      .catch(function (error) {
        log('Error:' + JSON.stringify(error));
      });
```
# Dev Machine
When test the addin on the dev machine, we could copy the manifest to dev catalog and the use `http-server --cors` to start the webside.

# Maintainers
[JackyChen200304](https://github.com/JackyChen200304)
