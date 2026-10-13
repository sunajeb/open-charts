// XML manifest of the Office.js add-in (PowerPoint only): a task pane opened from one ribbon button
// "open-charts" on the Home tab. PowerPointApi 1.4 is the floor (shapes, text, fill/line format);
// everything above (selection 1.5, groups / z-order / slide export 1.8, rotation / adjustments /
// theme colours 1.10) is feature-detected at run time (src/officejs/plan.ts hostCaps), so older
// PowerPoints still get the add-in and fall back to Path A.
import { ENGINE_VERSION } from '../version.ts';

/** Fixed add-in id: sideloaded copies, centralized deployment and AppSource all use the same one. */
export const ADDIN_ID = '0c8f4a52-6b1d-4e7a-9f3c-5d2e8a1b7c64';
export const DEFAULT_BASE = 'https://localhost:4175';

const esc = (s: string) => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

export function manifestXml(base = DEFAULT_BASE, version = `${ENGINE_VERSION}.0`): string {
  const u = (p: string) => esc(`${base.replace(/\/$/, '')}/${p}`);
  const icons = '<bt:Image size="16" resid="Icon.16"/><bt:Image size="32" resid="Icon.32"/><bt:Image size="80" resid="Icon.80"/>';
  return `<?xml version="1.0" encoding="UTF-8"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
  xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
  xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
  xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides"
  xsi:type="TaskPaneApp">
  <Id>${ADDIN_ID}</Id>
  <Version>${esc(version)}</Version>
  <ProviderName>open-charts</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="open-charts"/>
  <Description DefaultValue="Consulting-quality charts as native, editable PowerPoint shapes: insert, select to edit the data, re-layout after a resize."/>
  <IconUrl DefaultValue="${u('assets/icon-32.png')}"/>
  <HighResolutionIconUrl DefaultValue="${u('assets/icon-64.png')}"/>
  <SupportUrl DefaultValue="${u('support.html')}"/>
  <AppDomains>
    <AppDomain>${esc(base.replace(/\/$/, ''))}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Presentation"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="PowerPointApi" MinVersion="1.4"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="${u('taskpane.html?host=office')}"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Presentation">
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="OpenCharts.Group">
                <Label resid="Group.Label"/>
                <Icon>${icons}</Icon>
                <Control xsi:type="Button" id="OpenCharts.Pane">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>
                  <Icon>${icons}</Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>OpenChartsPane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="${u('assets/icon-16.png')}"/>
        <bt:Image id="Icon.32" DefaultValue="${u('assets/icon-32.png')}"/>
        <bt:Image id="Icon.80" DefaultValue="${u('assets/icon-80.png')}"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="${u('commands.html')}"/>
        <bt:Url id="Taskpane.Url" DefaultValue="${u('taskpane.html?host=office')}"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Group.Label" DefaultValue="open-charts"/>
        <bt:String id="Button.Label" DefaultValue="open-charts"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="Button.Tooltip" DefaultValue="Open the open-charts pane: insert a chart on this slide, or select one to edit its data."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
`;
}
