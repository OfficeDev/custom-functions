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
## Create Contextual Tabs related API
```js
// Show the shared runtime
await Office.addin.showAsTaskpane();

// Hide the shared runtime
await Office.addin.hide();

// Add event handler when the taskpane visibility mode is changed.
var handlerRemove = await Office.addin.onVisibilityModeChanged(function(args) {
    console.log('Visibility is changed to ' + args.visibilityMode)
});

// Remove the handler
await handlerRemove();
```

// To know the initial visibility mode.
Office.onReady(function(hostInfo) {
  if (hostInfo.addin) { // it works on desktop client. It will work on Excel online once the code is deployed.
    console.log(hostInfo.addin.visibilityMode);
  }
})

# Dev Machine
When test the addin on the dev machine, we could copy the manifest to dev catalog and the use `http-server --cors` to start the webside.

# Maintainers
[JackyChen200304](https://github.com/JackyChen200304)
