declare interface IListColumnsAutoFetchWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  WebUrlFieldLabel: string;
  ListTitleFieldLabel: string;
  ColumnFieldLabel:string;
}

declare module 'ListColumnsAutoFetchWebPartStrings' {
  const strings: IListColumnsAutoFetchWebPartStrings;
  export = strings;
}
