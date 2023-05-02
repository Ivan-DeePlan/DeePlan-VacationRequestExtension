declare interface IVacationRequestCommandSetStrings {
  Command1: string;
  Command2: string;
}

declare module 'VacationRequestCommandSetStrings' {
  const strings: IVacationRequestCommandSetStrings;
  export = strings;
}
