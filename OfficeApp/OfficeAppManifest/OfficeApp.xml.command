<?xml version="1.0" encoding="UTF-8"?>
<!--Created:cb85b80c-f585-40ff-8bfc-12ff4d0e34a9-->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6d595f8b-0289-47d5-8aa9-c6bc68f61e2a</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>[Provider name]</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="Word Add-in Angular Typescript" />
  <Description DefaultValue="Word Add-in written in Typescript with Angular"/>
  <!--Icon for your add-in. Used on installation screens and the add-ins dialog -->
  <IconUrl DefaultValue="https://i.imgur.com/oZFS95h.png" />

  <!--BeginTaskpaneMode integration. Office 2013 and any client that doesn't understand commands will use this section.
    This section will also be used if there are no VersionOverrides -->
  <Hosts>
    <Host Name="Document" />
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="~remoteAppUrl/app/index.html#/demo" />
  </DefaultSettings>
  <!--EndTaskpaneMode integration -->
  
  <Permissions>ReadWriteDocument</Permissions>
  
  <!--BeginAddinCommandsMode integration-->
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      !--Each host can have a different set of commands. Cool huh!? -->
      <!-- Workbook=Excel Document=Word Presentation=PowerPoint -->
      <!-- Make sure the hosts you override match the hosts declared in the top section of the manifest -->
      <Host xsi:type="Document">
        <!-- Form factor. Currently only DesktopFormFactor is supported. TabletFormFactor and PhoneFormFactor will be added in the future-->
        <DesktopFormFactor>

          <!--PrimaryCommandSurface==Main Office Ribbon-->
          <ExtensionPoint xsi:type="PrimaryCommandSurface">

            <!--Use OfficeTab to extend an existing Tab. Use CustomTab to create a new tab -->
            <!-- Documentation includes all the IDs currently tested to work -->
            <CustomTab id="Tab1">

              <!--Group. Ensure you provide a unique id. Recommendation for any IDs is to namespace using your company name-->
              <Group id="Microsoft.DX.WordAddInAngularTypeScript.Tab.Group">

                <!--Label for your group. resid must point to a ShortString resource -->
                <Label resid="Tab.GroupLabel" />
                <Icon>
                  <!-- Sample Todo: Each size needs its own icon resource or it will look distorted when resized -->
                  <!--Icons. Required sizes 16,31,80, optional 20, 24, 40, 48, 64. Strongly recommended to provide all sizes for great UX -->
                  <!--Use PNG icons and remember that all URLs on the resources section must use HTTPS -->
                  <bt:Image size="16" resid="Equipment.Icon" />
                  <bt:Image size="32" resid="Equipment.Icon" />
                  <bt:Image size="80" resid="Equipment.Icon" />
                </Icon>

            
                <Control xsi:type="Button" id="Taskpane1Button">
                  <Label resid="Taskpane1Button.Label" />
                  <Supertip>
                    <Title resid="Taskpane1Button.Label" />
                    <Description resid="Taskpane1Button.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Tools.Icon" />
                    <bt:Image size="32" resid="Tools.Icon" />
                    <bt:Image size="80" resid="Tools.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Taskpane1Button</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Taskpane1.Url" />
                  </Action>
                </Control>

 
                <Control xsi:type="Button" id="Taskpane2Button">
                  <Label resid="Taskpane2Button.Label" />
                  <Supertip>
                    <Title resid="Taskpane2Button.Label" />
                    <Description resid="Taskpane2Button.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Import.Icon" />
                    <bt:Image size="32" resid="Import.Icon" />
                    <bt:Image size="80" resid="import.Icon" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Taskpane2Button</TaskpaneId>
                    <!--Provide a url resource id for the location that will be displayed on the task pane -->
                    <SourceLocation resid="Taskpane2.Url" />
                  </Action>
                </Control>
              </Group>
              <!-- Label of your tab -->
              <!-- If validating with XSD it needs to be at the end - for now -->
              <Label resid="Tab.TabLabel" />
            </CustomTab>
          </ExtensionPoint>

        </DesktopFormFactor>
      </Host>
    </Hosts>
     <Resources>
      <bt:Images>
        <bt:Image id="Equipment.Icon" DefaultValue="https://localhost:44300/images/equipment.png" />
		    <bt:Image id="Import.Icon" DefaultValue="https://localhost:44300/images/import.png" />
        <bt:Image id="Tools.Icon" DefaultValue="https://localhost:44300/images/tools.png" />
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Taskpane1.Url" DefaultValue="https://localhost:4300/app/index#home" />
        <bt:Url id="Taskpane2.Url" DefaultValue="https://localhost:4300/app/index#demo" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Taskpane1Button.Label" DefaultValue="Insert Text Sample" />
        <bt:String id="Taskpane2Button.Label" DefaultValue="Replace Content Contols Sample" />
        <bt:String id="Tab.GroupLabel" DefaultValue="Angular Typescript Add-in Sample" />
        <bt:String id="Tab.TabLabel" DefaultValue="Add-in Sample" />
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Taskpane1Button.Tooltip" DefaultValue="Click to Show a Taskpane" />
        <bt:String id="Taskpane2Button.Tooltip" DefaultValue="Click to Show a Taskpane" />
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
