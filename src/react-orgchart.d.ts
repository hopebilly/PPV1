// src/types/react-orgchart.d.ts or types/react-orgchart.d.ts
declare module 'react-orgchart';

declare global {
    interface JQuery {
      orgchart(method: string, ...args: any[]): any;
      orgchart(options: any): JQuery;
    }
  }