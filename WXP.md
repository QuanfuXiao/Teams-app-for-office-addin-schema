# WXP features in extensions

## Scenarios
This doc introduces some properties in schema which are related to Teams app for Office add-in. (WXP devPreview) \
Generally the flowing features are enabled:
- Taskpane add-in
- Content add-in
- Context menu
- GetStartedMessage

## Schema details

### extension
| Property | Type | Required |Description |Notes                                       
| ---- | ---- | ---- | ------------------------------------------------------------- | ----|
| `runtimes`  | array of [extensionRuntime](#extensionruntime) |  | Genereral runtime is for "MailApp" or "TaskpaneApp", it configures the sets of runtimes and actions that can be used by each extension point. Min size 1. |
| `contentRuntimes`  | array of [extensionContentRuntime](#extensioncontentruntime) |  | Content runtime is for "ContentApp", which can be embedded directly into Excel or PowerPoint documents. Min size 1. |
| `getStartedMessages` | array of [extensionGetStartedMessage](#extensiongetstartedmessage) | | Provides information used by the callout that appears when the add-in is installed. Min size 1. Max size 3.|
| `contextMenus` | array of [extensionContextMenu](#extensioncontextmenu) |  | A context menu is a shortcut menu that appears when you right-click in the Office UI. Min size 1. | 


### extensionContentRuntime
This is the runtime for content add-in.
|Name|Type|Maximum size|Required|Description|
|:----|:----|:----|:----|:----|
|`id`|string|64 characters|✔️|A unique identifier for this runtime within the app.  This is developer specified.
|`code`|extensionRuntimeCode| |✔️|Specifies the location of code for this runtime. Depending on the `runtime.type`, add-ins use either a JavaScript file or an HTML page with an embedded `<script>` tag that specifies the URL of a JavaScript file.|
|`requestedWidth`|number| | |The desired width in pixels for the initial content placeholder. This value MUST be between 32 and 1000 pixels. Default value will be determined by host.||
|`requestedHeight`|number| | |The desired height in pixels for the initial content placeholder. This value MUST be between 32 and 1000 pixels. Default value will be determined by host.||
| `disableSnapshot`  | boolean | |  | Specifies whether a snapshot image of your content add-in is saved with the host document. Default value is `false`. Set `true` to disable.||
|`requirements`|[extensionRequirements](#extensionrequirements) | |  | Specifies the Office requirement sets for content add-in runtime. If the user's Office version doesn't support the specified requirements, the component will not be available in that client. ||

### Example:
```json
"contentRuntimes": [
    {
        "id": "ContentRuntime",
        "code": {
        "page": "https://localhost:3000/content.html"
        },
        "requestedWidth": 100,
        "requestedHeight": 100,
        "disableSnapshot": true,
        "requirements": {
            "scopes": [
                "workbook"
            ]
        },
    }
]
```

### extensionGetStartedMessage
Specifies the Get Started information for the Office add-in. This information is used at various places on the Office User Interface after user installs an add-in.
|Name|Type|Maximum size|Required|Description|
|:----|:----|:----|:----|:----|
|`requirements`|[extensionRequirements](#extensionrequirements) |||
|`title`|string|125 characters | ✔️|The title used for the top of the callout.|
|`description`|string|250 characters |✔️|	The description / body content for the callout.|
|`learnMoreUrl`|url|2048 characters |✔️ |	A URL to a page that explains the add-in in detail.|

### Example:
```json
"getStartedMessages": [
    {
        "requirements": {
        "scopes": [
            "workbook"
        ]
        },
        "title": "Get Started with Contoso Add-in",
        "description": "Opens a pane displaying all available properties.",
        "learnMoreUrl": "https://localhost:3000/learnMoreUrl.html"
    }
]
```

### extensionContextMenu
A context menu is a shortcut menu that appears when you right-click in the Office UI.
|Name|Type|Maximum size|Required|Description|
|:----|:----|:----|:----|:----|
|`requirements`|[extensionRequirements](#extensionrequirements) |||
|`menus`|array of [extensionMenu](#extensionmenu)| | ✔️|The title used for the top of the callout. Min size 1. |

### extensionMenu
|Name|Type|Maximum size|Required|Description|
|:----|:----|:----|:----|:----|
|`entryPoint`|[extensionMenuEntryPointEnum](#extensionmenuentrypointenum)||✔️|Use `text` or `cell` here for Office context menu. Use `text` if the context menu should open when a user right-clicks on the selected text. Use `cell` if the context menu should open when the user right-clicks on a cell on an Excel spreadsheet.
|`controls`|array of [extensionControl](#extensioncontrol)| | ✔️|The control type should be `menu`.Min size 1. |

### extensionMenuEntryPointEnum
|Name|Description|
|:----|:----|
|`text`|context menu should open when a user right-clicks on the selected text.|
|`cell`|context menu should open when the user right-clicks on a cell on an Excel spreadsheet.|

### Example:
```json
"contextMenus": [
    {
        "requirements": {
            "scopes": [
                "workbook"
            ],
            "capabilities": [
                {
                    "name": "AddinCommands",
                    "minVersion": "1.1"
                }
            ]
        },
        "menus": [
            {
                "entryPoint": "cell",
                "controls": [
                    {
                        "id": "menuForCell",
                        "type": "menu",
                        "label": "Menu",
                        "icons": [
                            {
                                "size": 16,
                                "url": "https://officedev.github.io/testing-assets/addins/images/button16x16.png"
                            },
                            {
                                "size": 32,
                                "url": "https://officedev.github.io/testing-assets/addins/images/button32x32.png"
                            },
                            {
                                "size": 80,
                                "url": "https://officedev.github.io/testing-assets/addins/images/button80x80.png"
                            }
                        ],
                        "supertip": {
                            "title": "Change text case",
                            "description": "This allow you to change text to lowercase or uppercase."
                        },
                        "items": [
                            {
                                "id": "menu.uppercase",
                                "type": "menuItem",
                                "label": "To uppercase",
                                "supertip": {
                                    "title": "Change text to uppercase",
                                    "description": "This will change the text to uppercase."
                                },
                                "actionId": "text.toUppercase"
                            }
                        ]
                    }
                ]
            }
        ]
    }
]
```
### extensionRequirements
Specify the capabilities, scopes and form factors that this extension should be available for and can run in.
|Name|Type|Maximum size|Required|Description|
|:----|:----|:----|:----|:----|
|`scopes`|array of [extensionRequirementsScopeEnum](#extensionrequirementsscopeenum)|min1, max 4| |Identifies the scopes in which the add-in can run. Support `mail`, `workbook`, `document`, `presentation`|

### extensionRequirementsScopeEnum
|Name|Description|
|:----|:----|
|`mail`|Outlook|
|`workbook`|Excel|
|`document`|Word|
|`presentation`|PowerPoint|

### extensionRuntime
Specifies a web or script execution environment/context.  Any extension taskpanes or script execution has to occur inside of such a context.
|Name|Type|Maximum size|Required|Description|
|:----|:----|:----|:----|:----|
|`id`|string|64 characters|✔️|A unique identifier for this runtime within the app.  This is developer specified.
|`type`|string enum| | |Specifies the type of runtime. Currently supports `general` for supporting running functions and launching pages.  [browser-based runtime](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes#browser-runtime).|
|`code`|[extensionRuntimeCode]| |✔️|Specifies the location of code for this runtime. Depending on the `runtime.type`, add-ins use either a JavaScript file or an HTML page with an embedded `<script>` tag that specifies the URL of a JavaScript file.|
|`lifetime`|string enum| | |Runtimes with a `short` lifetime do not preserve state across executions; runtimes with a `long` lifetime do. For more information about runtime lifetime, see [Runtimes in Office Add-ins](https://learn.microsoft.com/en-us/office/dev/add-ins/testing/runtimes).  Default value is `short`.|
|`actions`|array of [extensionRuntimeAction]| | |Specifies the set of actions supported by this runtime. An action is either running a JavaScript function or opening a view such as a task pane.  This is an array of objects, which are described in the extensions[0].runtime[0].actions[0] section below.|
| `requirements`  | [extensionRequirements](#extensionrequirements) | |  | Specifies the Office requirement sets for an add-in or add-in component like ribbon, autorun etc. If the user's Office version doesn't support the specified requirements, the extension or component will not be available in that client. Requirements are supported at the extension, alternates, ribbon, autoRun and runtime levels.  In each case, the service only returns or the Office client will only load items that match the host. ||

### extensionControl
Specifies a control used within a ribbon group or menu.
|Name|Type|Maximum size|Required|Description|
|:----|:----|:----|:----|:----|
|`actionId`|string|64 characters|✔️|Identifies the action that is taken when a user selects the control. The _actionId_ must be an exact match for a `runtime.actions.id`.|
|`builtinControlId`|string|64 characters|✔️|Id of the existing Office control. See [Find the IDs of controls and control groups](https://learn.microsoft.com/en-us/office/dev/add-ins/design/built-in-button-integration#find-the-ids-of-controls-and-control-groups).|
|`enabled`|boolean| | |Indicates whether the control is initially enabled. Default is `true`.|
|`icons`| array of [extensionCommonIcon]| |✔️|Configures the icons for the custom item.|
|`id`|string|64 characters|✔️|Unique identifier for this control within the app. Must be different from any built-in control id in the Office application and any other custom control.  This is developer specified.|
|`items`|array of [extensionMenuItem]| | |Configures the items for a menu control.|
|`label`|string|64 characters|✔️|Displayed text for the control.|
|`overriddenByRibbonApi`|boolean| | |Specifies whether a button, menu, or menu item will be hidden on application and platform combinations that support the API ([Office.ribbon.requestCreateControls](https://learn.microsoft.com/en-us/javascript/api/office/office.ribbon#office-office-ribbon-requestcreatecontrols-member(1))) that installs custom contextual tabs on the ribbon. Default is `false`.|
|`supertip`|[extensionSuperTip]| |✔️|Configures a supertip for the control.|
|`type`|string| |✔️|Supported values: `button`, `menu`.|

## Requested validations
#### 1. `extensioncontexts` should have corresponding `extensionScopes` 

| Context | Scope                                                          |
| ---- | ------------------------------------------------------------- |
| "mailRead","mailCompose","meetingDetailsOrganizer","meetingDetailsAttendee","onlineMeetingDetailsOrganizer" "logEventMeetingDetailsAttendee"   | `mail` |
| "default" | `workbook`, `document`, `presentation` |

#### 2. Either runtimes property or contentRuntimes property must be defined, but not both at the same time.  

## Localization string
"^extensions\\[[0-9]\\]\\.getStartedMessages\\[[0-2]\\]\\.title$" \
"^extensions\\[[0-9]\\]\\.getStartedMessages\\[[0-2]\\]\\.description$" \
"^extensions\\[[0-9]\\]\\.getStartedMessages\\[[0-2]\\]\\.learnMoreUrl$" \
"^extensions\\[[0-9]\\]\\.contentRuntimes\\[[1]?[0-9]\\]\\.code\\.page$" 

## Note
More features are on the way:
- Shortcut
- Custom function