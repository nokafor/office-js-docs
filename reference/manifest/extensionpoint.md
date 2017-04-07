# ExtensionPoint element

## Description

 Defines where an add-in integrates with the Office UI. 
 
 ### Example
 The following examples show how to use the  **ExtensionPoint** element with **PrimaryCommandSurface** and **ContextMenu** attribute values, and the child elements that should be used with each.


 >**Important**  For elements that contain an ID attribute, make sure you provide a unique ID. We recommend that you use your company's name along with your ID. For example, use the following format.<CustomTab id="mycompanyname.mygroupname">


```XML
<OfficeApp>
  ...
  <VersionOverrides>
   ...
   <ExtensionPoint xsi:type="PrimaryCommandSurface">
     <CustomTab id="Contoso Tab">
     <!-- If you want to use a default tab that comes with Office, remove the above CustomTab element, and then uncomment the following OfficeTab element -->
      <!-- <OfficeTab id="TabData"> -->
       <Label resid="residLabel4" />
       <Group id="Group1Id12">
         <Label resid="residLabel4" />
         <Icon>
           <bt:Image size="16" resid="icon1_32x32" />
           <bt:Image size="32" resid="icon1_32x32" />
           <bt:Image size="80" resid="icon1_32x32" />
         </Icon>
         <Tooltip resid="residToolTip" />
         <Control xsi:type="Button" id="Button1Id1">

            <!-- information about the control -->
         </Control>
         <!-- other controls, as needed -->
       </Group>
     </CustomTab>
   </ExtensionPoint>

  <ExtensionPoint xsi:type="ContextMenu">
   <OfficeMenu id="ContextMenuCell">
     <Control xsi:type="Menu" id="ContextMenu2">
            <!-- information about the control -->
     </Control>
    <!-- other controls, as needed -->
   </OfficeMenu>
  </ExtensionPoint>
  ...
 </VersionOverrides>
 ...
</OfficeApp>
```

## Attributes

|  Attribute  |  Required  |  Description  |
|:-----|:-----|:-----|
|  **xsi:type**  |  Yes  | The type of extension point being defined.|

### Extension point options for Word, Excel, PowerPoint, and OneNote add-in commands
|Extension point type | Description|
|-|-|
|PrimaryCommandSurface | Puts buttons on the ribbon in Office.|
|ContextMenu| Puts command on the shortcut menu that appears when you right-click in the Office UI.|

### Extension point options for Outlook add-in commands

|Extension point type | Description|
|-|-|
|[MessageReadCommandSurface](#messagereadcommandsurface) | Puts buttons in the command surface for the mail read view. In Outlook desktop, this appears in the ribbon.|
|[MessageComposeCommandSurface](#messagecomposecommandsurface) | Puts buttons on the ribbon for add-ins using mail compose form. |
|[AppointmentOrganizerCommandSurface](#appointmentorganizercommandsurface)| Puts buttons on the ribbon for the form that's displayed to the organizer of the meeting. | 
|[AppointmentAttendeeCommandSurface](#appointmentattendeecommandsurface)| Puts buttons on the ribbon for the form that's displayed to the attendee of the meeting. |
|[Module](#module) (Can only be used in the [DesktopFormFactor](./formfactor.md).) | Puts buttons on the ribbon for the module extension. |

## Child elements

The ExtensionPoint element must have at least one of the following child elements.
> Note: OfficeMenu is not a child element for Outlook extension points

_Does OfficeMenu only work for Word or Excel?_
 
|**Element**|**Description**|
|:-----|:-----|
|[CustomTab]()| Adds the command(s) to the custom ribbon tab. If you use the  **CustomTab** element, you can't use the **OfficeTab** element. The **id** attribute is required.|
|[OfficeTab]()|Adds the command(s) to the default ribbon tab. If you use the  **OfficeTab** element, you can't use the **CustomTab** element. For details, see [OfficeTab](officetab.md).|
|[OfficeMenu]()|Adds the command(s) to a default context menu. The  **id** attribute must be set to: <br/> - **ContextMenuText** for Excel or Word to display the item on the context menu when text is selected and then the user right-clicks on the selected text. <br/> - **ContextMenuCell** for Excel to display the  item on the context menu when the user right-clicks on a cell on the spreadsheet.|




## Parent Element

[FormFactor](./formfactor.md)

## Additional Information
None
