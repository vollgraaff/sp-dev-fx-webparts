import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { IAdaptiveCardActionResult } from "./IAdaptiveCardActionResult";

export interface IAdaptiveCardProps {
  themeVariant?: IReadonlyTheme | undefined;
  template: string;
  data: string;
  useTemplating: boolean;
  useArrayCycling: boolean;
  // currentIndex: number;
  className?: string;
  onExecuteAction?: (action: IAdaptiveCardActionResult) => void;
  onParseSuccess?: () => void;
  onParseError?: (errors: Array<string>) => void;
}
