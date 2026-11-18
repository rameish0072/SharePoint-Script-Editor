import { DisplayMode } from '@microsoft/sp-core-library';

export interface IScriptinjectorProps {
  htmlInput:       string;
  htmlFileName:    string;
  hasTeamsContext: boolean;
  isDarkTheme:     boolean;
  instanceId:      string;
  displayMode:     DisplayMode;
}