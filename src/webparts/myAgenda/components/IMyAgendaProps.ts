import { DisplayMode } from "@microsoft/sp-core-library";
export interface IMyAgendaProps {
  description: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}
