import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IAdaptiveCardHostProps {
  themeVariant: IReadonlyTheme | undefined;
  template: string;
  data: string;
  useTemplating: boolean;
  useArrayCycling: boolean;
  displayMode: DisplayMode;
  context: WebPartContext;
}

export interface IAdaptiveCardHostState {
  currentIndex: number;
  dataLength: number;
}
