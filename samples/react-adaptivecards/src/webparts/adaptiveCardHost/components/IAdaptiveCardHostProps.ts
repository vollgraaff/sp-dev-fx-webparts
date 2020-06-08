import { DisplayMode } from "@microsoft/sp-core-library";
import { WebPartContext, BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";

export interface IAdaptiveCardHostProps {
  themeVariant: IReadonlyTheme | undefined;
  template: string;
  data: string;
  useTemplating: boolean;
  useArrayCycling: boolean;
  displayMode: DisplayMode;
  context: WebPartContext;
  // WebpartElement: BaseClientSideWebPart<IAdaptiveCardHostProps>;
}

export interface IAdaptiveCardHostState {
  currentIndex: number;
  dataLength: number;
  // WindowSize: number;
}
