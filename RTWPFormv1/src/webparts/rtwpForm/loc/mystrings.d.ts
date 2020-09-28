declare interface IRtwpFormWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  WorkFieldLabel:string;
  WorkCFieldLabel:string;
  BuildingFieldLabel:string;
}

declare module 'RtwpFormWebPartStrings' {
  const strings: IRtwpFormWebPartStrings;
  export = strings;
}
